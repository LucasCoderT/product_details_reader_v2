class NoSkuColumnsFound(Exception):
    """Exception raised when no SKU columns are found in the DataFrame."""
    def __init__(self, file_name: str, columns: list[str] = None):
        super().__init__(f"No SKU columns found in the file: {file_name}, please check the file. Found the following columns: {columns}")
        self.file_name = file_name
        self.columns = []