class person:
    def __init__(self, name, age):
        self.name = name
        self.age = age

    def printname(self):
        print(f"Employee name is {self.name}, age is {self.age}")
        if self.age >= 30:
            print("old")
        
a = person("hannan", 23)

a.printname()