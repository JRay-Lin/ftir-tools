"""
File Conversion Module for FTIR Spectroscopy

Contains functions for converting JASCO .jws files to YLK format
and managing YLK data structures.
"""

import os
import pandas as pd
import olefile
import struct
import json
from datetime import datetime


def data_definitions(x):
    """Define data type mappings for JASCO files"""
    return {
        268435715: "WAVELENGTH",
        4097: "CD",
        8193: "HT VOLTAGE",
        3: "ABSORBANCE",
        14: "FLUORESCENCE",
    }.get(x, "undefined")


class JwsHeader:
    """JWS file header class"""

    def __init__(
        self,
        channel_number,
        point_number,
        x_for_first_point,
        x_for_last_point,
        x_increment,
        header_names,
        data_size=None,
    ):
        self.channel_number = channel_number
        self.point_number = point_number
        self.x_for_first_point = x_for_first_point
        self.x_for_last_point = x_for_last_point
        self.x_increment = x_increment
        self.data_size = data_size
        self.header_names = header_names


def _unpack_ole_jws_header(data):
    """Parse OLE JWS file header"""
    try:
        data_tuple = struct.unpack("<LLLLLLddd", data[0:48])
        channels = data_tuple[3]
        nxtfmt = "<L" + "L" * channels
        header_names = list(struct.unpack(nxtfmt, data[48 : 48 + 4 * (channels + 1)]))

        for i, e in enumerate(header_names):
            header_names[i] = data_definitions(e)

        data_tuple += tuple(header_names)

        lastPos = 48 + 4 * (channels + 1)
        nxtfmt = "<LLdddd"
        for pos in range(channels):
            data_tuple = data_tuple + struct.unpack(
                nxtfmt, data[lastPos : lastPos + 40]
            )
            lastPos += 40

        return JwsHeader(
            data_tuple[3],
            data_tuple[5],
            data_tuple[6],
            data_tuple[7],
            data_tuple[8],
            header_names,
        )
    except Exception as e:
        raise Exception(f"Unable to read DataInfo: {str(e)}")


def convert_jws_to_ylk_direct(jws_filename, ylk_filename):
    """
    Direct JWS file to YLK conversion (without depending on jws2txt)

    Parameters:
    jws_filename: path to input .jws file
    ylk_filename: path to output .ylk file

    Returns:
    bool: True if successful, False otherwise
    """
    try:
        with open(jws_filename, "rb") as f:
            oleobj = olefile.OleFileIO(f)

            # Read header information
            data_stream = oleobj.openstream("DataInfo")
            header_data = data_stream.read()
            header_obj = _unpack_ole_jws_header(header_data)

            # Read Y data
            fmt = "f" * header_obj.point_number * header_obj.channel_number
            y_data_stream = oleobj.openstream("Y-Data")
            values = struct.unpack(fmt, y_data_stream.read())

            chunks = [
                values[x : x + header_obj.point_number]
                for x in range(0, len(values), header_obj.point_number)
            ]

            oleobj.close()

        # Prepare data arrays
        x_data = []
        y_data = []

        for line_no in range(header_obj.point_number):
            x_value = header_obj.x_for_first_point + line_no * header_obj.x_increment
            y_value = chunks[0][line_no] if chunks else 0
            x_data.append(round(float(x_value), 6))
            y_data.append(round(float(y_value), 8))

        # Calculate range and round to nearest 10x
        x_min_raw, x_max_raw = min(x_data), max(x_data)

        # Round to nearest 10x (e.g., 1999.8 -> 2000, 2038 -> 2040)
        x_min = round(x_min_raw / 10) * 10
        x_max = round(x_max_raw / 10) * 10

        # Create YLK structure
        ylk_data = {
            "name": os.path.basename(jws_filename).replace(".jws", ""),
            "range": [x_min, x_max],
            "raw_data": {"x": x_data, "y": y_data},
            "baseline": {"x": [], "y": []},
            "metadata": {
                "created": datetime.now().isoformat(),
                "source_file": os.path.basename(jws_filename),
                "channels": header_obj.channel_number,
                "points": header_obj.point_number,
            },
        }

        # Write YLK file
        with open(ylk_filename, "w", encoding="utf-8") as f:
            json.dump(ylk_data, f, indent=2, ensure_ascii=False)

        return True

    except Exception as e:
        print(f"JWS file conversion failed {jws_filename}: {str(e)}")
        return False


def convert_jws_with_fallback(jws_file, output_dir):
    """
    Convert JWS file to YLK format using direct conversion

    Parameters:
    jws_file: path to input .jws file
    output_dir: output directory for YLK file

    Returns:
    str: path to converted YLK file, or None if failed
    """
    base_name = os.path.basename(jws_file).replace(".jws", "")
    ylk_file = os.path.join(output_dir, f"{base_name}.ylk")

    # Use direct conversion to YLK format
    if convert_jws_to_ylk_direct(jws_file, ylk_file):
        return ylk_file
    else:
        print(f"Unable to convert JWS file {os.path.basename(jws_file)}")
        return None


def load_ylk_file(ylk_filename):
    """
    Load YLK file and return data structure

    Parameters:
    ylk_filename: path to YLK file

    Returns:
    dict: YLK data structure, or None if failed
    """
    try:
        with open(ylk_filename, "r", encoding="utf-8") as f:
            data = json.load(f)
        return data
    except Exception as e:
        print(f"Unable to load YLK file {ylk_filename}: {str(e)}")
        return None


def save_ylk_file(ylk_filename, data):
    """
    Save data structure to YLK file

    Parameters:
    ylk_filename: path to YLK file
    data: YLK data structure to save

    Returns:
    bool: True if successful, False otherwise
    """
    try:
        # Update modification timestamp
        if "metadata" not in data:
            data["metadata"] = {}
        data["metadata"]["modified"] = datetime.now().isoformat()

        with open(ylk_filename, "w", encoding="utf-8") as f:
            json.dump(data, f, indent=2, ensure_ascii=False)
        return True
    except Exception as e:
        print(f"Unable to save YLK file {ylk_filename}: {str(e)}")
        return False


def ylk_to_dataframe(ylk_data):
    """
    Convert YLK data to pandas DataFrame for analysis

    Parameters:
    ylk_data: YLK data structure

    Returns:
    pd.DataFrame: DataFrame with wavenumber and absorbance columns
    """
    try:
        raw_data = ylk_data.get("raw_data", {})
        x_data = raw_data.get("x", [])
        y_data = raw_data.get("y", [])

        df = pd.DataFrame({"wavenumber": x_data, "absorbance": y_data})
        return df
    except Exception as e:
        print(f"Unable to convert YLK to DataFrame: {str(e)}")
        return None


def get_supported_formats():
    """
    Get list of supported file formats

    Returns:
    dict: format descriptions and extensions
    """
    return {
        "jws": "JASCO JWS files (binary format)",
        "ylk": "YLK data files (JSON format with baseline)",
    }
