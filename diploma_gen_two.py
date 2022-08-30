"""Create diploms for Lucky School."""
import win32com.client
import os

psApp = win32com.client.Dispatch("Photoshop.Application")

psApp.Open(r"F:\path\project.psd")  # Turn PH on

diploma = psApp.Application.ActiveDocument


class EditDiploma:
    """Edit diploma.psd data."""

    def __init__(self, owner_name, dog_name, reg_code, dog_breed, dog_class, number):
        """."""
        self.owner_name = owner_name
        self.dog_name = dog_name
        self.reg_code = reg_code
        self.dog_breed = dog_breed
        self.dog_class = dog_class
        self.number = number

        self.export_folder = "F:/output_path"

    def edit_owner_name(self):
        """."""
        owner_name_layer = diploma.ArtLayers["owner-name"]
        owner_text = owner_name_layer.TextItem
        owner_text.contents = self.owner_name

    def edit_dog_name(self):
        """."""
        dog_name_layer = diploma.ArtLayers["dog-name"]
        dog_text = dog_name_layer.TextItem
        dog_text.contents = self.dog_name

    def edit_reg_code(self):
        """."""
        date_layer = diploma.ArtLayers["reg-code"]
        date_text = date_layer.TextItem
        date_text.contents = self.reg_code

    def edit_dog_breed(self):
        """."""
        date_layer = diploma.ArtLayers["breed"]
        date_text = date_layer.TextItem

        if self.dog_breed == "LR":
            self.dog_breed = "Labradori Retriiver / Labrador Retriever"
        elif self.dog_breed == "GR":
            self.dog_breed = "Kuldne Retriiver / Golden Retriever"
        elif self.dog_breed == "FCR":
            self.dog_breed = "Siledakarvaline Retriiver / Flatcoated Retriever"
        elif self.dog_breed == "CBR":
            self.dog_breed = "Chesapeake Bay Retriiver / Chesapeake Bay Retriever"
        elif self.dog_breed == "NSDTR":
            self.dog_breed = "Nova Scotia Retriiver / Nova Scotia Duck Tolling Retriever"

        date_text.contents = self.dog_breed

    def edit_dog_class(self):
        """."""
        date_layer = diploma.ArtLayers["class"]
        date_text = date_layer.TextItem
        date_text.contents = self.dog_class

    def export_diploma(self):
        """."""
        options = win32com.client.Dispatch('Photoshop.ExportOptionsSaveForWeb')
        options.Format = 6  # JPEG
        options.Quality = 100  # Value from 0-100
        filename = f"{self.export_folder}\{self.number}.jpg"
        diploma.Export(ExportIn=filename, ExportAs=2, Options=options)


def read_names(filename):
    """."""
    data = []
    with open(filename, "r", encoding='utf-8') as csvfile:
        for line in csvfile:
            dog_class, name, dog, breed, reg = line.strip().split(",")
            person_data = {"owner_name": name.title(), "dog_name": dog.title(), "breed": breed, "reg": reg,
                           "dog_class": dog_class}
            data.append(person_data)
    return data


def create_diploma(data):
    """."""
    number = 1
    for diploma in data:
        owner_name = diploma['owner_name']
        dog_name = diploma['dog_name']
        breed = diploma['breed']
        reg = diploma['reg']
        dog_class = diploma['dog_class']
        diplom = EditDiploma(owner_name, dog_name, reg, breed, dog_class, number)
        diplom.edit_dog_name()
        diplom.edit_owner_name()
        diplom.edit_dog_class()
        diplom.edit_reg_code()
        diplom.edit_dog_breed()
        diplom.export_diploma()
        number += 1
    return True


data = read_names("data_file.csv")
print(data)
create_diploma(data)
