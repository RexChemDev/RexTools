import json
import pandas as pd
import opentrons.execute
from string import ascii_uppercase

def custom_labware(file_name, location):
    protocol = opentrons.execute.get_protocol_api("2.12")
    with open(str(file_name)) as labware_file:
        labware_def = json.load(labware_file)
        return protocol.load_labware_from_definition(labware_def, int(location))


def coord_iter(last_letter="H", last_number=12):
    letter_sequence = ascii_uppercase[:ascii_uppercase.index(last_letter) + 1]
    number_sequence = [i for i in range(1, last_number + 1)]
    coord_grid = [[f"{letter}{number}" for number in number_sequence] for letter in letter_sequence]
    for row in coord_grid:
        for item in row:
            yield item

class Worker:
    
    def __init__(self, sheet_name, index_col=0, vol_col=12, last_letter="H", last_number=12):
        self.index_col, self.vol_col = index_col, vol_col
        self.raw_instructions = self._parse_excel(sheet_name)
        self.sequence = list(coord_iter(last_letter, last_number))
        self.commands = list(map(self._command, self.raw_instructions))
        self._index = -1
    
    @property
    def is_finished(self):
        if self._index >= len(self.commands) - 1:
            return True
        else:
            return False


    def _command(self, sheet_entry: list):
        try:
            attempt_index = sheet_entry[self.index_col] - 1
            location = self.sequence[attempt_index]
        except IndexError:
            print("WARNING: There are more compounds than plate positions.")
            print(f"Command translation has stopped at index {attempt_index}.")
            raise StopIteration # cuts off mapping prematurely.
        return {
            "location": location,
            "volume": int(sheet_entry[self.vol_col])
        }
    
    def _parse_excel(self, sheet_name):
        print("Reading spreadsheet data...")
        xl_data = pd.read_excel(sheet_name, engine="openpyxl")
        raw_instructions = xl_data.values.tolist()
        print("Finished reading.")
        return raw_instructions
    
    def __iter__(self):
        return self
    
    def __next__(self):
        if self.is_finished:
            raise StopIteration
        self._index += 1
        return self.commands[self._index]

        


__all__ = [
    "custom_labware",
    "coord_iter",
    "Worker",
]