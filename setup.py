from setuptools import setup, find_packages

setup(
    name="excel_template_app",
    version="0.1.1",
    packages=find_packages(where="src"),
    package_dir={"": "src"},
    install_requires=[
        "pandas",
        "xlwings",
        "python-docx",
    ],
    entry_points={
        "console_scripts": [
            "excel_template_app=excel_template_app.app:main",
        ],
    },
    author="Larry Grullon-Polanco",
    author_email="larrygrpolanco@gmail.com",
    description="An application for applying Excel templates to transcript data for a research project.",
    long_description=open("README.md").read(),
    long_description_content_type="text/markdown",
    url="https://github.com/larrygrpolanco/transcript-excel-converter",
    classifiers=[
        "Programming Language :: Python :: 3",
        "License :: OSI Approved :: MIT License",
        "Operating System :: OS Independent",
    ],
    python_requires=">=3.6",
)
