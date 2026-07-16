import urllib.parse


def canonicalize_path(path: str) -> str:
    """Normalize Explorer locations for case-insensitive comparison.

    Local ``file:///C:/...`` URLs and UNC ``file://server/share/...`` URLs
    need different prefix handling. Parsing the URL also ensures percent
    escapes are decoded consistently for both forms.
    """
    if not path:
        return ""
    path = path.strip()
    if path.lower().startswith("file:"):
        parsed = urllib.parse.urlsplit(path)
        decoded_path = urllib.parse.unquote(parsed.path)
        if parsed.netloc and parsed.netloc.lower() != "localhost":
            path = f"//{urllib.parse.unquote(parsed.netloc)}{decoded_path}"
        else:
            # urlsplit preserves the slash before a Windows drive letter.
            if len(decoded_path) >= 3 and decoded_path[0] == "/" and decoded_path[2] == ":":
                decoded_path = decoded_path[1:]
            path = decoded_path
    path = path.replace("/", "\\")
    while len(path) > 3 and path.endswith("\\"):
        path = path[:-1]
    return path.lower()
