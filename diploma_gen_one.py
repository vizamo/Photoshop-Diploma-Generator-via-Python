import win32com.client
import os

psApp = win32com.client.Dispatch("Photoshop.Application")
psApp.Open(r"F:\path\project.psd")
badge_psd = psApp.Application.ActiveDocument


def read_badge_list(file="cup-osalejad-list.csv"):
    persons = []
    with open(file, "r", encoding="utf-8") as csv_file:
        for line in csv_file:
            osaleja = line.strip().split(",")
            o_class = osaleja[0]
            o_name = osaleja[1]
            o_dog = osaleja[2]
            person_data = {"o_class": o_class, "o_name": o_name.title(), "o_dog": o_dog.title()}
            persons.append(person_data)
    return persons


def generator(persons_data):
    x = 1
    y = 1
    for person in persons_data:
        person_badge = DiplomsGenerator(person['o_class'], person['o_name'], person['o_dog'], x)
        person_badge.edit_class()
        person_badge.edit_name()
        person_badge.edit_dog()
        if x % 2 == 0:
            save_badge(y)
            y += 1
            x = 0
        x += 1
    print("Job is done")


def save_badge(y):
    options = win32com.client.Dispatch('Photoshop.ExportOptionsSaveForWeb')
    options.Format = 6  # JPEG
    options.Quality = 100  # Value from 0-100
    path = "D:F:/output_path"
    filename = f"{path}/{y}.jpg"
    badge_psd.Export(ExportIn=filename, ExportAs=2, Options=options)

class DiplomsGenerator:

    def __init__(self, o_class, o_name, o_dog, number):
        self.o_class = o_class
        self.o_name = o_name
        self.o_dog = o_dog
        self.number = number

    def edit_class(self):
        """."""
        layer = f"class_{self.number}"
        f_name_layer = badge_psd.LayerSets[self.number - 1].ArtLayers[layer]
        f_name_text = f_name_layer.TextItem
        f_name_text.contents = self.o_class

    def edit_name(self):
        """."""
        layer = f"name_{self.number}"
        l_name_layer = badge_psd.LayerSets[self.number - 1].ArtLayers[layer]
        l_name_text = l_name_layer.TextItem
        l_name_text.contents = self.o_name

    def edit_dog(self):
        """."""
        layer = f"dog_{self.number}"
        role_layer = badge_psd.LayerSets[self.number - 1].ArtLayers[layer]
        role_text = role_layer.TextItem
        role_text.contents = self.o_dog


data = read_badge_list()
generator(data)
print(data)
