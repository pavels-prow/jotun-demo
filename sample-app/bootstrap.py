import os
import subprocess
import sys
import venv


def venv_python_path(venv_dir: str) -> str:
    if os.name == "nt":
        return os.path.join(venv_dir, "Scripts", "python.exe")
    return os.path.join(venv_dir, "bin", "python")


def ensure_venv(venv_dir: str) -> None:
    if os.path.isdir(venv_dir):
        return
    builder = venv.EnvBuilder(with_pip=True, clear=False)
    builder.create(venv_dir)


def run() -> int:
    project_dir = os.path.dirname(os.path.abspath(__file__))
    venv_dir = os.path.join(project_dir, ".venv")
    ensure_venv(venv_dir)
    python_exe = venv_python_path(venv_dir)

    requirements = os.path.join(project_dir, "requirements.txt")
    subprocess.check_call([python_exe, "-m", "pip", "install", "-r", requirements])
    return subprocess.call([python_exe, os.path.join(project_dir, "app.py")])


if __name__ == "__main__":
    raise SystemExit(run())
