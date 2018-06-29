import setuptools

with open("README.md", "r") as fh:
    long_description = fh.read()

setuptools.setup(
    name="femapy",
    version="1.0",
    author="Andy Perez",
    long_description=long_description,
    long_description_content_type="text/markdown",
    packages=setuptools.find_packages()
)