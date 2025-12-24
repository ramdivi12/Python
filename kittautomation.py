#!/usr/bin/env python3

import os
import argparse
import subprocess
import logging
import json
import shutil
from copy import deepcopy
from ruamel.yaml import YAML
from ruamel.yaml.comments import CommentedMap, CommentedSeq

# ================= YAML =================
yaml = YAML()
yaml.preserve_quotes = True
yaml.indent(mapping=2, sequence=4, offset=2)

# ================= Logging =================
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s"
)
logger = logging.getLogger(__name__)

# ================= Globals =================
BACKUP_DIR = ".kitt-backup"
CHANGE_REPORT = []
CREATED_FILES = []

# ================= Helpers =================
def ensure_dir(path):
    os.makedirs(path, exist_ok=True)

def backup_file(repo, file_path):
    rel = os.path.relpath(file_path, repo)
    backup_path = os.path.join(repo, BACKUP_DIR, rel + ".bak")
    ensure_dir(os.path.dirname(backup_path))
    if not os.path.exists(backup_path):
        shutil.copy2(file_path, backup_path)

def record_change(file, field, old, new):
    CHANGE_REPORT.append({
        "file": file,
        "field": field,
        "old": old,
        "new": new
    })

# ================= IO =================
def load_yaml(path):
    with open(path, "r") as f:
        return yaml.load(f) or CommentedMap()

def write_yaml(path, data, repo, dry_run):
    if dry_run:
        logger.info(f"[DRY-RUN] Would write {path}")
        return
    backup_file(repo, path)
    with open(path, "w") as f:
        yaml.dump(data, f)

# ================= Fixes =================
def fix_cluster_id(node, cluster_id, file):
    if isinstance(node, dict):
        for k, v in node.items():
            if k == "cluster_id":
                old = v
                seq = CommentedSeq([str(cluster_id)])
                seq.fa.set_flow_style()  # INLINE LIST
                node[k] = seq
                record_change(file, "cluster_id", old, seq)
            else:
                fix_cluster_id(v, cluster_id, file)
    elif isinstance(node, list):
        for i in node:
            fix_cluster_id(i, cluster_id, file)

def fix_namespace(node, namespace, file):
    if isinstance(node, dict):
        for k, v in node.items():
            if k == "namespace":
                old = v
                node[k] = namespace
                record_change(file, "namespace", old, namespace)
            else:
                fix_namespace(v, namespace, file)
    elif isinstance(node, list):
        for i in node:
            fix_namespace(i, namespace, file)

def fix_country(node, country, file):
    if isinstance(node, dict):
        for k, v in node.items():
            if k == "labels" and isinstance(v, dict) and "ccm.country" in v:
                old = v["ccm.country"]
                v["ccm.country"] = country
                record_change(file, "ccm.country", old, country)
            fix_country(v, country, file)
    elif isinstance(node, list):
        for i in node:
            fix_country(i, country, file)

def fix_cnames(node, country, file):
    if isinstance(node, dict):
        for k, v in node.items():
            if k == "cnames" and isinstance(v, list):
                for idx, c in enumerate(v):
                    if "cell000" in c:
                        old = c
                        v[idx] = c.replace("cell000", country)
                        record_change(file, "cnames", old, v[idx])
            else:
                fix_cnames(v, country, file)
    elif isinstance(node, list):
        for i in node:
            fix_cnames(i, country, file)

# ================= Templates =================
def discover_templates(files):
    return [
        f for f in files
        if f.startswith("kitt") and f.endswith(".yml")
        and (
            "us-wm.primary" in f
            or "us-wm.secondary" in f
            or f in ("kitt.primary.yml", "kitt.secondary.yml")
            or "cell000" in f
        )
    ]

# ================= Create country files =================
def create_country_files(path, countries, cluster_id, namespace, repo, dry_run):
    created = []
    for tpl in discover_templates(os.listdir(path)):
        src = os.path.join(path, tpl)
        base = load_yaml(src)

        for country in countries:
            dest = tpl.replace("cell000", country).replace("us-wm", country)
            if dest == tpl:
                dest = f"kitt.{country}.{tpl.replace('kitt.', '')}"

            dest_path = os.path.join(path, dest)
            if os.path.exists(dest_path):
                continue

            data = deepcopy(base)
            fix_cluster_id(data, cluster_id, dest)
            fix_namespace(data, namespace, dest)
            fix_country(data, country, dest)
            fix_cnames(data, country, dest)

            if not dry_run:
                ensure_dir(os.path.dirname(dest_path))
                with open(dest_path, "w") as f:
                    yaml.dump(data, f)

            #CREATED_FILES.append(dest_path)
            CREATED_FILES.append(os.path.relpath(dest_path, repo))
            created.append(dest)
            logger.info(f"Created {dest}")

    return created

# ================= Pipeline =================
def update_pipeline(path, created, repo, dry_run):
    kitt = os.path.join(path, "kitt.yml")
    if not os.path.exists(kitt):
        return

    data = load_yaml(kitt)
    post = data.get("build", {}).get("postBuild")
    if not isinstance(post, list):
        return

    for f in created:
        post.append(CommentedMap({
            "task": CommentedMap({
                "name": "deployApp",
                "kittFilePath": f"{os.path.basename(path)}/{f}",
                "sync": False,
                "executionScope": "child"
            })
        }))

    write_yaml(kitt, data, repo, dry_run)

# ================= Traverse =================
def traverse_repo(repo, countries, cluster_id, namespace, dry_run):
    for root, _, files in os.walk(repo):
        if any(f.startswith("kitt") for f in files):
            created = create_country_files(
                root, countries, cluster_id, namespace, repo, dry_run
            )
            if created:
                update_pipeline(root, created, repo, dry_run)


def persist_created_files(repo):
    if not CREATED_FILES:
        return
    path = os.path.join(repo, BACKUP_DIR, "created-files.json")
    ensure_dir(os.path.dirname(path))
    with open(path, "w") as f:
        json.dump(CREATED_FILES, f, indent=2)


# ================= Rollback =================
def rollback(repo):
    backup_root = os.path.join(repo, BACKUP_DIR)
    if not os.path.exists(backup_root):
        logger.error("No backup directory found")
        return

    # 1. Restore backups
    for root, _, files in os.walk(backup_root):
        for f in files:
            if f.endswith(".bak"):
                backup_file = os.path.join(root, f)
                original = os.path.join(
                    repo,
                    os.path.relpath(backup_file, backup_root)[:-4]
                )
                ensure_dir(os.path.dirname(original))
                shutil.copy2(backup_file, original)
                logger.info(f"Restored {original}")

    # 2. Delete newly created files
    created_list = os.path.join(backup_root, "created-files.json")
    if os.path.exists(created_list):
        with open(created_list) as f:
            created_files = json.load(f)

        for rel_path in created_files:
            abs_path = os.path.join(repo, rel_path)
            if os.path.exists(abs_path):
                os.remove(abs_path)
                logger.info(f"Deleted created file {abs_path}")

    logger.info("Rollback completed successfully")


# ================= CLI =================
def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--repo", required=True)
    ap.add_argument("--countries")
    ap.add_argument("--cluster-id")
    ap.add_argument("--namespace")
    ap.add_argument("--dry-run", action="store_true")
    ap.add_argument("--rollback", action="store_true")
    ap.add_argument("--report", default="kitt-change-report.json")
    args = ap.parse_args()

    if args.rollback:
        rollback(args.repo)
        return

    traverse_repo(
        args.repo,
        [c.strip() for c in args.countries.split(",")],
        args.cluster_id,
        args.namespace,
        args.dry_run
    )

    persist_created_files(args.repo)

    if CHANGE_REPORT:
        with open(args.report, "w") as f:
            json.dump(CHANGE_REPORT, f, indent=2)
        logger.info(f"Change report written to {args.report}")

    if not args.dry_run:
        subprocess.run(["git", "diff"], cwd=args.repo)

    logger.info("Completed successfully")

if __name__ == "__main__":
    main()
