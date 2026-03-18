from setuptools import find_packages, setup

setup(
    name="iia-excel-reorg",
    version="0.1.0",
    description="Workflow to reorganize historical Excel workbooks into a standardized structure.",
    packages=find_packages(where="workflow/src"),
    package_dir={"": "workflow/src"},
    python_requires=">=3.11",
    extras_require={"dev": ["pytest>=8.0.0"]},
    entry_points={"console_scripts": ["iia-excel-reorg=iia_excel_reorg.cli:main"]},
)
