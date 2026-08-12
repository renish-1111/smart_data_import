from setuptools import setup, find_packages

setup(
    name="smart_data_import",
    version="0.0.1",
    description="High-Performance Multi-File Data Import Engine with Automated DAG Dependency Sorting for ERPNext",
    author="ERPNext AI Team",
    author_email="info@erpnext.ai",
    packages=find_packages(),
    zip_safe=False,
    include_package_data=True,
    install_requires=["openpyxl"],
)
