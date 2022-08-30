import win32com.client
import os

psApp = win32com.client.Dispatch("Photoshop.Application")
psApp.Open(r"F:\path\project.psd")
badge_psd = psApp.Application.ActiveDocument


def read_badge_list(file="badge_list.txt"):
    persons = []
    with open(file, "r", encoding="utf") as txt_file:
        for line in txt_file:
            full_name, role = line.strip().split(", ")
            name = full_name.split(" ")
            first_name = " ".join(name[:-1])
            last_name = name[-1]

            person_data = {"f_name": first_name.capitalize(), "l_name": last_name.capitalize(), "role": role}
            persons.append(person_data)
    return persons


def generator(persons_data):
    x = 1
    y = 1
    for person in persons_data:
        person_badge = BadgeGenerator(person['f_name'], person['l_name'], person['role'], x)
        person_badge.edit_first_name()
        person_badge.edit_last_name()
        person_badge.edit_role()
        if x % 10 == 0:
            save_badge(y)
            y += 1
            x = 0
        x += 1
    if x != 1:
        while x != 11:
            visible_folder = badge_psd.LayerSets[x - 1]
            visible_folder.visible = False
            x += 1
        save_badge(y)
    print("Job is done")


def save_badge(y):
    options = win32com.client.Dispatch('Photoshop.ExportOptionsSaveForWeb')
    options.Format = 2
    options.PNG8 = False
    path = "F:/output_path/"
    filename = f"{path}/badge_list_{y}.png"
    badge_psd.Export(ExportIn=filename, ExportAs=2, Options=options)


class BadgeGenerator:

    def __init__(self, first_name, last_name, role, number):
        self.first_name = first_name
        self.last_name = last_name
        self.role = role
        self.number = number

        visible_folder = badge_psd.LayerSets[self.number - 1]
        visible_folder.visible = True

    def edit_first_name(self):
        """."""
        layer = f"first_name_{self.number}"
        f_name_layer = badge_psd.LayerSets[self.number - 1].ArtLayers[layer]
        f_name_text = f_name_layer.TextItem
        f_name_text.contents = self.first_name

    def edit_last_name(self):
        """."""
        layer = f"last_name_{self.number}"
        l_name_layer = badge_psd.LayerSets[self.number - 1].ArtLayers[layer]
        l_name_text = l_name_layer.TextItem
        l_name_text.contents = self.last_name

    def edit_role(self):
        """."""
        layer = f"role_{self.number}"
        role_layer = badge_psd.LayerSets[self.number - 1].ArtLayers[layer]
        role_text = role_layer.TextItem
        role_text.contents = self.role


data = read_badge_list()
generator(data)
