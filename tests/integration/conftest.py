def pytest_configure(config):
    config.addinivalue_line(
        "markers", "requires_qt: marks tests that require an initialized PySide6/Qt runtime"
    )
