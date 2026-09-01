"""The schema, entity and BREX resources bundled with this package.

`XmlChecker` and `BrexChecker` both resolve a reference the same way: look
where the checked object is first, then in whatever the caller registered,
then here. "Here" is the package's own resource tree, shipped as package
data so a checkout or an install needs no configuration at all::

    acd/
      xsd/2.3/  xsd/2.3.1/  xsd/4.0/  xsd/4.0.1/  xsd/4.0.2/
      xsd/4.1/  xsd/4.2/    xsd/5.0/  xsd/6/          -- S1000D schemas by issue
      ent/catalog.xml  ent/4.1/  ent/4.2/  ...        -- ISO character entities
      brex/                                           -- S1000D master BREX
      brex/brex_2_3/  brex_4_2/  brex_5_0/  brex_6_0/

An application may also ship its own `res/xsd` and `res/brex` tree beside
itself -- an ALTHOM application does -- and that is searched *after* the
bundled copies, so a project can add an issue or a schema this package does
not carry without replacing anything. `find_resource_root` locates it by
walking out from the running script; `ACD_RES_DIR` names one outright.

`catalog.xml` is the offline answer to the ISO entity sets S1000D content
references by URL (`http://www.s1000d.org/S1000D_4-1/ent/...`). Registering
it with libxml2 is what lets a data module carrying `&ndash;` parse on a
machine with no network; `register_catalog` does that, and both checkers call
it for you.

Nothing here raises. A resource tree that is absent simply contributes no
directories, and the caller reports the reference as unresolved.
"""

from os import environ
from os import listdir
from os.path import abspath
from os.path import basename
from os.path import dirname
from os.path import isdir
from os.path import isfile
from os.path import join
from sys import argv
from sys import modules
from sys import prefix
from urllib.parse import urlparse

from .default_brex import DEFAULT_BREX_DIR

# This package's own resource directories, shipped as package data.
PACKAGE_DIR = dirname(abspath(__file__))
BUNDLED_XSD_DIR = join(PACKAGE_DIR, "xsd")
BUNDLED_ENT_DIR = join(PACKAGE_DIR, "ent")
BUNDLED_BREX_DIR = DEFAULT_BREX_DIR
BUNDLED_CATALOG = join(BUNDLED_ENT_DIR, "catalog.xml")

# Environment variable naming an *additional* resource tree, searched after
# the bundled one. Set it when an application keeps its `res` somewhere the
# walk below will not find.
RES_DIR_ENV = "ACD_RES_DIR"

# How far up from a starting directory to look for `res`. Deep enough to
# climb out of `res/py/scripts/authoring/` (4) with room to spare, shallow
# enough never to wander into the drive root.
_MAX_WALK_UP = 8

# Cached result of `find_resource_root`. Discovery touches the filesystem a
# dozen times; a folder run must not repeat that per object.
_ROOT_CACHE = {}


def _subdirectories(path: str) -> list:
    """The immediate subdirectories of `path`, sorted, or an empty list."""
    if not isdir(path):
        return []
    return [join(path, name) for name in sorted(listdir(path))
            if isdir(join(path, name))]


def _looks_like_resource_root(path: str) -> bool:
    """Whether `path` is a `res` tree rather than a directory that happens to
    be called res: it has to actually carry schemas or BREX modules."""
    return isdir(path) and (isdir(join(path, "xsd")) or isdir(join(path, "brex")))


def _walk_up_for_res(start: str) -> str:
    """Look for a `res` directory at `start` and in each parent above it.

    Args:
        start (str): directory to start from

    Returns:
        str: the resource root, or None
    """
    current = abspath(start)
    for _ in range(_MAX_WALK_UP):
        candidate = join(current, "res")
        if _looks_like_resource_root(candidate):
            return candidate
        # A bundle may be laid out with the tree directly at the root.
        if _looks_like_resource_root(current):
            return current
        parent = dirname(current)
        if parent == current:
            break
        current = parent
    return None


def find_resource_root(refresh: bool = False) -> str:
    """Locate an application's own `res/` tree, if it ships one.

    This is the *secondary* source: the bundled `acd/xsd` and `acd/brex` are
    consulted first, and this exists so a project can carry an issue or a
    project BREX the package does not.

    Search order:

    1. `ACD_RES_DIR`, if set and it looks like a resource tree.
    2. The directory of the running script, and its parents.
    3. A PyInstaller bundle's `_MEIPASS` directory, and its parents.
    4. The working directory and its parents.
    5. `sys.prefix`, for an application installed into a virtual environment.

    Args:
        refresh (bool): re-run discovery instead of reusing the cached answer

    Returns:
        str: absolute path of the resource root, or None
    """
    if not refresh and "root" in _ROOT_CACHE:
        return _ROOT_CACHE["root"]

    named = environ.get(RES_DIR_ENV)
    if named and _looks_like_resource_root(named):
        _ROOT_CACHE["root"] = abspath(named)
        return _ROOT_CACHE["root"]

    starts = []
    main_module = modules.get("__main__")
    main_file = getattr(main_module, "__file__", None)
    if main_file:
        starts.append(dirname(abspath(main_file)))
    if argv and argv[0]:
        starts.append(dirname(abspath(argv[0])))
    meipass = getattr(modules.get("sys"), "_MEIPASS", None) or environ.get("_MEIPASS2")
    if meipass:
        starts.append(meipass)
    starts.append(abspath("."))
    starts.append(prefix)

    root = None
    seen = set()
    for start in starts:
        if not start or start in seen or not isdir(start):
            continue
        seen.add(start)
        root = _walk_up_for_res(start)
        if root:
            break

    _ROOT_CACHE["root"] = root
    return root


# ---------------------------------------------------------------------------
# Schemas
# ---------------------------------------------------------------------------

def _version_token(declared: str) -> str:
    """The S1000D issue a schema URL names, in `xsd/` directory form.

    `http://www.s1000d.org/S1000D_4-1/xml_schema_flat/proced.xsd` -> `"4.1"`.

    Args:
        declared (str): the `xsi:noNamespaceSchemaLocation` value

    Returns:
        str: the version token, or None if the location does not name one
    """
    for part in urlparse(declared).path.split("/") + declared.split("/"):
        if part.upper().startswith("S1000D_"):
            return part[len("S1000D_"):].replace("-", ".")
    return None


def schema_dirs(root: str = None) -> list:
    """Every schema directory available, bundled first.

    Args:
        root (str): application resource root; discovered when omitted

    Returns:
        list: absolute directory paths -- `acd/xsd/<issue>/` followed by any
            `res/xsd/<issue>/` an application ships
    """
    directories = _subdirectories(BUNDLED_XSD_DIR)
    root = root if root is not None else find_resource_root()
    if root:
        directories.extend(_subdirectories(join(root, "xsd")))
    return directories


def find_schema(declared: str, root: str = None) -> str:
    """Find a declared schema in the bundled tree, then an application's own.

    The issue the location names is preferred, then progressively shorter
    forms of it (`4.0.2` falls back to `4.0`, and `6.0` to `6`), and finally
    any issue directory that has a file of that name. The last step matters
    because a CSDB is not always consistent about which issue's URL it quotes,
    and validating against a near neighbour beats not validating at all --
    which is why the checker records `source` on every resolved schema.

    Args:
        declared (str): the `xsi:noNamespaceSchemaLocation` value
        root (str): application resource root; discovered when omitted

    Returns:
        str: absolute path of the schema file, or None
    """
    name = basename(urlparse(declared).path) or basename(declared)
    if not name:
        return None
    directories = schema_dirs(root)
    if not directories:
        return None

    version = _version_token(declared)
    if version:
        # First match wins, and `schema_dirs` puts the bundled copies first,
        # so an application's tree only supplies what this package does not.
        by_name = {}
        for directory in directories:
            by_name.setdefault(basename(directory), directory)
        candidate = version
        while candidate:
            directory = by_name.get(candidate)
            if directory and isfile(join(directory, name)):
                return join(directory, name)
            if "." not in candidate:
                break
            candidate = candidate.rsplit(".", 1)[0]

    for directory in directories:
        candidate = join(directory, name)
        if isfile(candidate):
            return candidate
    return None


# ---------------------------------------------------------------------------
# BREX
# ---------------------------------------------------------------------------

def brex_dirs(root: str = None, include_default: bool = True) -> list:
    """Every directory that may hold a BREX data module.

    Bundled `acd/brex/<issue>/` folders first, then any `res/brex/<issue>/`
    an application ships, then optionally `acd/brex` itself.

    `BrexChecker` passes `include_default=False` and consults these only
    *after* `default_brex.find_default_brex_fallback` has had its turn. That
    path matches on the logical DMC, tolerates a different issue or inWork
    number, and records the substitution so the report can say a built-in
    BREX stood in for the referenced one. The `brex/<issue>/` folders hold
    copies of those same master modules -- `brex_2_3` carries an older issue
    of one of them -- so resolving them by a plain directory search first
    would answer the reference silently and lose the report entry.

    Args:
        root (str): application resource root; discovered when omitted
        include_default (bool): also return the flat bundled BREX directory

    Returns:
        list: absolute directory paths, in search order
    """
    directories = _subdirectories(BUNDLED_BREX_DIR)
    root = root if root is not None else find_resource_root()
    if root:
        brex_root = join(root, "brex")
        if isdir(brex_root):
            directories.extend(_subdirectories(brex_root))
            if any(_.lower().endswith(".xml") for _ in listdir(brex_root)):
                directories.append(brex_root)
    if include_default and isdir(BUNDLED_BREX_DIR):
        directories.append(BUNDLED_BREX_DIR)
    return directories


# ---------------------------------------------------------------------------
# Entities
# ---------------------------------------------------------------------------

def catalog_path(root: str = None) -> str:
    """The XML catalog mapping S1000D entity URLs to local files.

    The bundled `acd/ent/catalog.xml` is preferred; an application's
    `res/ent/catalog.xml` is the fallback. Its `rewritePrefix` values are
    relative, so the catalog only resolves from the directory it ships in --
    which is why this returns a path rather than reading it.

    Args:
        root (str): application resource root; discovered when omitted

    Returns:
        str: absolute path of the catalog, or None if there is not one
    """
    if isfile(BUNDLED_CATALOG):
        return BUNDLED_CATALOG
    root = root if root is not None else find_resource_root()
    if root:
        candidate = join(root, "ent", "catalog.xml")
        if isfile(candidate):
            return candidate
    return None


def register_catalog(path: str = None) -> str:
    """Register the entity catalog with libxml2, so entity references in
    S1000D content resolve without network access.

    S1000D data modules pull their ISO character entities from
    `http://www.s1000d.org/S1000D_<issue>/ent/...`. Without a catalog libxml2
    tries to fetch that; with one it reads the copy in `acd/ent/<issue>/`.

    lxml exposes no catalog-loading binding, so this appends to
    `XML_CATALOG_FILES`, which libxml2 reads the first time it consults the
    global catalog. That happens once per process, which is why both checkers
    call this when they are constructed rather than when they parse.

    Args:
        path (str): catalog to register; the bundled one when omitted

    Returns:
        str: the registered path, or None if there was no catalog to register
    """
    path = path or catalog_path()
    if not path or not isfile(path):
        return None
    entries = environ.get("XML_CATALOG_FILES", "").split()
    if path not in entries:
        entries.append(path)
        environ["XML_CATALOG_FILES"] = " ".join(entries)
    return path
