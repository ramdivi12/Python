import os
import sys
import shutil
import subprocess
import logging
from ruamel.yaml import YAML
from ruamel.yaml.comments import CommentedMap, CommentedSeq

# -----------------------
# Logging setup
# -----------------------
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s | %(levelname)s | %(message)s"
)
logger = logging.getLogger(__name__)

yaml = YAML()
yaml.preserve_quotes = True
yaml.indent(mapping=2, sequence=4, offset=2)

# -----------------------
# Helpers
# -----------------------
def load_yaml(path):
    with open(path, "r", encoding="utf-8") as f:
        return yaml.load(f)

def write_yaml(path, data):
    with open(path, "w", encoding="utf-8") as f:
        yaml.dump(data, f)

def is_git_repo(path):
    return os.path.isdir(os.path.join(path, ".git"))

def show_git_diff(path):
    if not is_git_repo(path):
        logger.warning("Not a git repository. Skipping git diff.")
        return
    logger.info("Showing git diff summary")
    subprocess.run(["git", "diff"], cwd=path)

# -----------------------
# Core logic
# -----------------------
def replace_cluster_and_namespace(data, cluster_id, namespace):
    """
    Safe replacement:
    - target can be map OR list of maps
    - cluster_id must remain a list
    """
    deploy = data.get("deploy")
    if not isinstance(deploy, CommentedMap):
        return

    stages = deploy.get("stages", [])
    if not isinstance(stages, (list, CommentedSeq)):
        return

    for stage in stages:
        if not isinstance(stage, CommentedMap):
            continue

        target = stage.get("target")

        # CASE 1: target is a MAP
        if isinstance(target, CommentedMap):
            if isinstance(target.get("cluster_id"), list) and target["cluster_id"]:
                target["cluster_id"][0] = cluster_id

        # CASE 2: target is a LIST of MAPS
        elif isinstance(target, CommentedSeq):
            for item in target:
                if isinstance(item, CommentedMap):
                    if isinstance(item.get("cluster_id"), list) and item["cluster_id"]:
                        item["cluster_id"][0] = cluster_id

        # Namespace replacement (helm.values.namespace)
        helm = stage.get("helm")
        if isinstance(helm, CommentedMap):
            values = helm.get("values")
            if isinstance(values, CommentedMap) and "namespace" in values:
                values["namespace"] = namespace


def copy_and_update_primary_secondary(module_dir, cluster_id, namespace):
    created_files = []

    primary_created = False
    secondary_created = False

    for fname in sorted(os.listdir(module_dir)):
        lower = fname.lower()

        # ---------- PRIMARY ----------
        if (
            lower.endswith(".primary.yml")
            and "intl" not in lower
            and not primary_created
        ):
            src = os.path.join(module_dir, fname)
            dest = os.path.join(module_dir, "Kirt.intl.primary.yml")

            if not os.path.exists(dest):
                logger.info(f"Creating intl primary from {fname}")
                shutil.copyfile(src, dest)

                data = load_yaml(dest)
                replace_cluster_and_namespace(data, cluster_id, namespace)
                write_yaml(dest, data)

                created_files.append("Kirt.intl.primary.yml")

            primary_created = True

        # ---------- SECONDARY ----------
        if (
            lower.endswith(".secondary.yml")
            and "intl" not in lower
            and not secondary_created
        ):
            src = os.path.join(module_dir, fname)
            dest = os.path.join(module_dir, "Kirt.intl.secondary.yml")

            if not os.path.exists(dest):
                logger.info(f"Creating intl secondary from {fname}")
                shutil.copyfile(src, dest)

                data = load_yaml(dest)
                replace_cluster_and_namespace(data, cluster_id, namespace)
                write_yaml(dest, data)

                created_files.append("Kirt.intl.secondary.yml")

            secondary_created = True

        # Exit early if both are done
        if primary_created and secondary_created:
            break

    return created_files


def append_tasks_to_kitt_yml(kitt_path, module_name, created_files):
    """
    Append deployApp tasks only once.
    """
    if not created_files:
        return

    data = load_yaml(kitt_path)

    build = data.setdefault("build", CommentedMap())
    post_build = build.setdefault("postBuild", CommentedSeq())

    existing_paths = set()

    for task in post_build:
        if isinstance(task, CommentedMap):
            path = task.get("kittFilePath")
            if path:
                existing_paths.add(path)

    for fname in created_files:
        kitt_path_ref = f"{module_name}/{fname}"
        if kitt_path_ref in existing_paths:
            logger.info(f"Task already exists, skipping: {kitt_path_ref}")
            continue

        logger.info(f"Appending task for: {kitt_path_ref}")
        task = CommentedMap({
            "task": CommentedMap({
                "name": "deployApp",
                "kittFilePath": kitt_path_ref,
                "sha": "{{$.kitt.build.commitEvent.commitId}}",
                "branch": "{{$.kitt.build.commitEvent.currentBranch}}",
                "sync": False,
                "executionScope": "child"
            })
        })
        post_build.append(task)

    write_yaml(kitt_path, data)


def process_module(module_dir, project_root, cluster_id, namespace):
    kitt_path = os.path.join(module_dir, "kitt.yml")
    if not os.path.exists(kitt_path):
        return

    logger.info(f"Processing module: {module_dir}")

    created = copy_and_update_primary_secondary(
        module_dir,
        cluster_id,
        namespace
    )

    if created:
        module_name = os.path.relpath(module_dir, project_root).replace("\\", "/")
        append_tasks_to_kitt_yml(kitt_path, module_name, created)


# -----------------------
# Main
# -----------------------
def main():
    project_root = input("Enter project repo root path: ").strip()
    if not os.path.isdir(project_root):
        logger.error("Invalid project root path")
        sys.exit(1)

    cluster_id = input("Enter intl cluster id: ").strip()
    namespace = input("Enter intl namespace: ").strip()

    logger.info("Starting processing...")

    for root, dirs, files in os.walk(project_root):
        if "kitt.yml" in files:
            process_module(root, project_root, cluster_id, namespace)

    logger.info("Processing completed")
    show_git_diff(project_root)


if __name__ == "__main__":
    main()
