from PyQt5.QtWidgets import QWidget
class ClassStorage:
    """
    # Пример использования
    storage = ClassStorage()
    storage.add_value("answer", 42)
    storage.add_value("pi", 3.14)
    storage.add_value("greeting", "hello")
    storage.add_value("another_answer", 100)
    storage.add_value("farewell", "world")

    print(storage)  # Выводит все хранимые значения, отсортированные по классам
    print(storage.get_values_by_class(int))  # Получает все значения типа int
    print(storage.get_values_by_class(str))  # Получает все значения типа str
    print(storage.get_values_by_class(float))  # Получает все значения типа float
    print(storage.get_value("pi"))  # Получает значение по названию "pi"
    """

    def __init__(self):
        self.storage = {}
        self.names = {}
        self.widgets_names = []

    def add_value(self, name, value):
        value_class = type(value)
        if value_class not in self.storage:
            self.storage[value_class] = {}
        self.storage[value_class][name] = value
        self.names[name] = value_class
        self.widgets_names.append(name)

    def get_values_by_class(self, cls):
        return self.storage.get(cls, {})

    def get_value(self, name):
        value_class = self.names.get(name)
        if value_class:
            return self.storage[value_class].get(name)
        return None
    
    def to_json(self):
        json_data = {}
        for cls, values in self.storage.items():
            class_name = cls.__name__
            json_data[class_name] = {}
            for name, value in values.items():
                if isinstance(value, QWidget):
                    if hasattr(value, 'text'):
                        json_data[class_name][name] = value.text()
                    elif hasattr(value, 'currentText'):
                        json_data[class_name][name] = value.currentText()
                    elif hasattr(value, 'isChecked'):
                        json_data[class_name][name] = value.isChecked()
                    elif hasattr(value, 'selectedDate'):
                        json_data[class_name][name] = value.selectedDate().toString("yyyy-MM-dd")
                    elif hasattr(value, 'toPlainText'):
                        json_data[class_name][name] = value.toPlainText()
        return json_data

    def __repr__(self):
        return f"{self.__class__.__name__}({self.storage})"
    
    
# Пример использования
storage = ClassStorage()
storage.add_value("answer", 42)
storage.add_value("pi", 3.14)
storage.add_value("greeting", "hello")
storage.add_value("another_answer", 100)
storage.add_value("farewell", "world")

print(storage)  # Выводит все хранимые значения, отсортированные по классам
print(storage.get_values_by_class(int))  # Получает все значения типа int
print(storage.get_values_by_class(str))  # Получает все значения типа str
print(storage.get_values_by_class(float))  # Получает все значения типа float
print(storage.get_value("pi"))  # Получает значение по названию "pi"