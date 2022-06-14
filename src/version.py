import re


__version__ = '0.1.1.dev0'


def format_version(version: str) -> str:
    """Extracts only the x.y.z semantic version tag."""

    try:
        return re.match(r'\d+.\d+.\d+', version).group(0)
    except IndexError:
        raise ValueError(f'Incorrect version given "{version}".') from None


def is_dev_version(version: str) -> str:
    """Denotes if the given version is dev."""
    
    return 'dev' in version
