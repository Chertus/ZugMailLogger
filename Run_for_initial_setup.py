import subprocess
import sys

def install_packages(requirements_path):
    try:
        subprocess.check_call([sys.executable, "-m", "pip", "install", "-r", requirements_path])
        print("All packages installed successfully.")
    except subprocess.CalledProcessError as e:
        print(f"An error occurred while installing packages: {e}")
        sys.exit(1)

if __name__ == "__main__":
    requirements_path = 'requirements.txt'  # Path to your requirements.txt file
    install_packages(requirements_path)
