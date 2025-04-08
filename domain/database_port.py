from abc import ABC, abstractmethod


class DatabasePort(ABC):
    @abstractmethod
    def connect(self):
        pass

    @abstractmethod
    def disconnect(Self):
        pass

    @abstractmethod
    def execute_query(self, query, params=None):
        pass
