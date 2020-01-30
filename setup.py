import os

from setuptools import setup, find_packages

here = os.path.abspath(os.path.dirname(__file__))
with open(os.path.join(here, 'README.md')) as f:
    README = f.read()
with open(os.path.join(here, 'CHANGES.txt')) as f:
    CHANGES = f.read()

requires = [
    'numpy',
    'jdcal',
    'openpyxl',
    ]

tests_require = []


def read_build_version():
    from configparser import ConfigParser

    DEFAULT_VERSION = '0.0.1'

    build_file = os.path.join(here, 'build.properties')
    if os.path.exists(build_file):
        config = ConfigParser()
        with open(build_file) as lines:
            lines = '[top]\n' + lines.read()
            config.read_string(lines)
            if 'version' in config['top']:
                return config['top']['version']

    return DEFAULT_VERSION


setup(name='excel-numpy',
      version=read_build_version(),
      description='excel-numpy',
      long_description=README + '\n\n' + CHANGES,
      classifiers=[
          "Programming Language :: Python",
      ],
      author='',
      author_email='',
      url='',
      keywords='excel numpy parser',
      packages=find_packages(),
      include_package_data=True,
      zip_safe=False,
      extras_require={
          'testing': tests_require,
      },
      install_requires=requires,
      entry_points="""\
      [paste.app_factory]
      main = server:main
      """,
      )
