import glob


from input_sheet import InputSheet
from output_gsheet import GsheetConverter

WINDOWS_GLOB = "C:\\Users\\yusai\\Downloads\\Q*.xlsx"
MAC_GLOB = "/Users/yusairak/Downloads//Q*.xlsx"

if __name__ == "__main__":
    records = []
    for file in glob.glob(MAC_GLOB):
        print("using file", file)
        i = InputSheet(file).load()
        records.append(i.output)
    records.sort(key=lambda x:x.title(), reverse=True)
    GsheetConverter.write(records)



