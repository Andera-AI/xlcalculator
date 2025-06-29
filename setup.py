import os

import setuptools


def read(*rnames):
    return open(os.path.join(os.path.dirname(__file__), *rnames)).read()


TESTS_REQUIRE = [
    "coverage",
    "flake8",
    "mock",
    "pytest",
    "pytest-cov",
]


setuptools.setup(
    name="xlcalculator",
    version="0.5.1.dev0",
    author="Bradley van Ree",
    author_email="brads@bradbase.net",
    description="Converts MS Excel formulas to Python and evaluates them.",
    long_description=(read("README.rst") + "\n\n" + read("CHANGES.rst")),
    url="https://github.com/bradbase/xlcalculator",
    packages=setuptools.find_packages(),
    license="MIT",
    keywords=[
        "xls",
        "Excel",
        "spreadsheet",
        "workbook",
        "data analysis",
        "analysis" "reading excel",
        "excel formula",
        "excel formulas",
        "excel equations",
        "excel equation",
        "formula",
        "formulas",
        "equation",
        "equations",
        "timeseries",
        "time series",
        "research",
        "scenario analysis",
        "scenario",
        "modelling",
        "model",
        "unit testing",
        "testing",
        "audit",
        "calculation",
        "evaluation",
        "data science",
        "openpyxl",
    ],
    classifiers=[
        "Development Status :: 5 - Production/Stable",
        "Intended Audience :: Developers",
        "Intended Audience :: Science/Research",
        "Intended Audience :: Information Technology",
        "Intended Audience :: Financial and Insurance Industry",
        "License :: OSI Approved :: MIT License",
        "Natural Language :: English",
        "Operating System :: OS Independent",
        "Programming Language :: Python :: 3",
        "Programming Language :: Python :: 3.9",
        "Programming Language :: Python :: 3.10",
        "Programming Language :: Python :: 3.11",
        "Topic :: Office/Business :: Financial :: Spreadsheet",
        "Topic :: Scientific/Engineering",
        "Topic :: Scientific/Engineering :: Information Analysis",
        "Topic :: Scientific/Engineering :: Mathematics",
        "Topic :: Software Development :: Libraries",
        "Topic :: Software Development :: Libraries :: Python Modules",
        "Topic :: Software Development :: Testing",
        "Topic :: Software Development :: Testing :: Unit",
        "Topic :: Utilities",
    ],
    install_requires=[
        "jsonpickle",
        "numpy>=1.20.3",
        "pandas>=2.0.0",
        "openpyxl",
        "numpy-financial",
        "mock",
        "scipy>1.10.0",
    ],
    dependency_links=[
        "git+https://github.com/Andera-AI/yearfrac.git#egg=yearfrac",
    ],
    extras_require=dict(
        test=TESTS_REQUIRE,
        build=[
            "pip-tools",
        ],
    ),
    python_requires=">=3.9",
    tests_require=TESTS_REQUIRE,
    include_package_data=True,
    zip_safe=False,
)
