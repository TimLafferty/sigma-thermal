"""
Sigma Thermal - Industrial Heater Design and Calculation Library
Setup configuration
"""

from setuptools import setup, find_packages
from pathlib import Path

# Read the README file
this_directory = Path(__file__).parent
long_description = (this_directory / "README.md").read_text() if (this_directory / "README.md").exists() else ""

# Read requirements
def read_requirements(filename):
    with open(filename) as f:
        return [line.strip() for line in f if line.strip() and not line.startswith('#') and not line.startswith('-r')]

setup(
    name="sigma-thermal",
    version="0.1.0",
    author="GTS Energy Inc",
    author_email="",
    description="Industrial heater design and calculation library",
    long_description=long_description,
    long_description_content_type="text/markdown",
    url="https://github.com/gts-energy/sigma-thermal",
    project_urls={
        "Bug Tracker": "https://github.com/gts-energy/sigma-thermal/issues",
        "Documentation": "https://sigma-thermal.readthedocs.io",
    },
    classifiers=[
        "Development Status :: 3 - Alpha",
        "Intended Audience :: Science/Research",
        "Intended Audience :: Manufacturing",
        "Topic :: Scientific/Engineering",
        "Topic :: Scientific/Engineering :: Chemistry",
        "License :: OSI Approved :: MIT License",
        "Programming Language :: Python :: 3",
        "Programming Language :: Python :: 3.11",
        "Programming Language :: Python :: 3.12",
    ],
    package_dir={"": "src"},
    packages=find_packages(where="src"),
    python_requires=">=3.11",
    install_requires=read_requirements("requirements.txt"),
    extras_require={
        "dev": read_requirements("requirements-dev.txt"),
    },
    entry_points={
        "console_scripts": [
            "sigma-thermal=sigma_thermal.cli:main",
        ],
    },
    include_package_data=True,
    package_data={
        "sigma_thermal": [
            "data/**/*.json",
            "data/**/*.csv",
            "reporting/templates/**/*.docx",
        ],
    },
)
