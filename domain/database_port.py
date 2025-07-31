from abc import ABC, abstractmethod

class DatabasePort(ABC):
    @abstractmethod
    def connect(self):
        pass

    @abstractmethod
    def disconnect(self):
        pass

    @abstractmethod
    def execute_query(self, query, params=None):
        pass

    @abstractmethod
    def execute_non_query(self, query, params=None):
        pass

    @abstractmethod
    def start_transaction(self):
        pass

    @abstractmethod
    def commit(self):
        pass

    @abstractmethod
    def rollback(self):
        pass