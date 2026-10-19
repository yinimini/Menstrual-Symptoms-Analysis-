from extraction.json_reader import read_json
import pandas as pd
import json

class DataProcessor:

    def __init__(self, data: dict) -> None:
        """
        Initialize the processor with the source data.

        Args:
            data (dict): Original dictionary containing the nested data to process.

        """
        self.data = data

    def extract_dataframe(
        self,
        parent_key: str,
        child_key: str
    ) -> pd.DataFrame:
        """
        Extract nested records and create a DataFrame.

        Args:
            parent_key (str): Key in the original dictionary whose value contains the child key.
            child_key (str): Key in the nested dictionary to extract.

        Returns:
            pd.DataFrame: A DataFrame containing the extracted records.
        """
        return pd.DataFrame(
            self.data[parent_key][child_key]
        )

    def flatten_user_patterns_data(
        self,
        df: pd.DataFrame
    ) -> pd.DataFrame:
        """
        Parse and flatten JSON-encoded user pattern data.

        Parses the JSON strings in the 'data' column, expands the
        resulting dictionaries into separate columns, and combines
        them with the original columns excluding 'data'.

        Args:
            df (pd.DataFrame): The DataFrame containing the JSON-encoded user pattern data.

        Returns:
            pd.DataFrame: The flattened DataFrame containing user pattern data.
        """
        df['data'] = df["data"].apply(json.loads)

        df_user_patterns = pd.concat(
            [
                df.drop(columns=["data"]),
                pd.json_normalize(df['data'])
            ],
            axis=1
        )

        return df_user_patterns


if __name__ == "__main__":
    data = read_json()