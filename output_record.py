import datetime



from lserm_row import LsermRow


class OutputRecord:


    def __init__(self, metadata=None):
        self._metadata = metadata
        self._all_current = []
        self._etfs_eligible = []
        self._all_eligible = []
        self._etfs_added = []
        self._all_deleted = []
        self._etfs_deleted = []

    def record_current(self, row: LsermRow):
        self.__record__(row, self._all_current, self._etfs_eligible)

    def record_addition(self, row: LsermRow):
        self.__record__(row, self._all_eligible, self._etfs_added)

    def record_deletion(self, row: LsermRow):
        self.__record__(row, self._all_deleted, self._etfs_deleted)



    def print(self):
        print("Metadata", self._metadata)
        print("Current", len(self._all_current), len(self._etfs_eligible))
        self.print_table(self._all_current, self._etfs_eligible)
        print("Added", len(self._all_eligible), len(self._etfs_added))
        self.print_table(self._all_eligible, self._etfs_added)
        print("Deleted", len(self._all_deleted), len(self._etfs_deleted))
        self.print_table(self._all_deleted, self._etfs_deleted)

    @staticmethod
    def print_table(list1, list2):
        for i in list1:
            print(i, "(ETF)" if i in list2 else "")

    @staticmethod
    def __record__(row: LsermRow, all_list: list[LsermRow], etf_list: list[LsermRow]):
        all_list.append(row)
        if row.is_etf():
            etf_list.append(row)
