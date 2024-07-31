import subprocess
import sys


class PackageInstaller:
    def __init__(self, packages):
        self.packages = packages

    def install_package(self, package_name):
        try:
            subprocess.check_call([sys.executable, "-m", "pip", "install", package_name])
            print(f"{package_name} installed successfully.")
        except Exception as e:
            print(f"Error installing {package_name}: {e}")

    def install_packages(self):
        for package in self.packages:
            try:
                __import__(package)
                print(f"{package} is already installed.")
            except ImportError:
                print(f"{package} not installed. Installing...")
                self.install_package(package)
