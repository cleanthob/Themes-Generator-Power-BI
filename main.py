import json
import os
import re
from tkinter import messagebox, ttk
from tkinter.filedialog import askopenfilename, asksaveasfilename

from ttkthemes import ThemedTk


class ThemesGenerator:
    def __init__(self):
        self._power_bi_fonts = [
            "Arial",
            "Arial Black",
            "Comic Sans MS",
            "Courier New",
            "DIN",
            "Georgia",
            "Impact",
            "Lucida Console",
            "Lucida Sans Unicode",
            "Microsoft Sans Serif",
            "Segoe UI",
            "Tahoma",
            "Times New Roman",
            "Trebuchet MS",
            "Verdana",
        ]
        self._show_error = False
        self._data = None
        self._root = ThemedTk(theme="breeze")
        self._width = 800
        self._height = 300
        self._center_window()
        self._root.resizable(False, False)
        self._root.title("Themes Generator")
        self._frame_configurations = ttk.Frame(self._root, width=200)
        self._label_theme_name = ttk.Label(
            self._frame_configurations, text="Nome do tema:"
        )
        self._label_theme_name.grid(column=0, row=0)
        self._entry_theme_name = ttk.Entry(self._frame_configurations)
        self._entry_theme_name.grid(column=1, row=0)
        self._label_font_size_value_card = ttk.Label(
            self._frame_configurations, text="Tamanho dos valores dos cartões:"
        )
        self._label_font_size_value_card.grid(column=0, row=1)
        self._entry_font_size_values_card = ttk.Entry(self._frame_configurations)
        self._entry_font_size_values_card.grid(column=1, row=1)
        self._label_font_color = ttk.Label(
            self._frame_configurations, text="Cor dos títulos dos visuais:"
        )
        self._label_font_color.grid(column=0, row=2)
        self._entry_font_color = ttk.Entry(self._frame_configurations)
        self._entry_font_color.grid(column=1, row=2)
        self._label_font = ttk.Label(
            self._frame_configurations, text="Fonte dos títulos dos visuais:"
        )
        self._label_font.grid(column=0, row=3)
        self._combobox_font = ttk.Combobox(
            self._frame_configurations,
            values=self._power_bi_fonts,
            width=18,
            state="readonly",
        )
        self._combobox_font.grid(column=1, row=3)
        self._label_background_color = ttk.Label(
            self._frame_configurations, text="Cor do fundo dos visuais:"
        )
        self._label_background_color.grid(column=0, row=4)
        self._entry_background_color = ttk.Entry(self._frame_configurations)
        self._entry_background_color.grid(column=1, row=4)
        self._btn_generate_theme = ttk.Button(
            self._frame_configurations, text="Gerar tema", command=self._generate_theme
        )
        self._btn_load_theme = ttk.Button(
            self._frame_configurations, text="Carregar tema", command=self._load_theme
        )
        self._btn_load_theme.grid(column=0, columnspan=2, row=5, pady=3)
        self._btn_generate_theme.grid(column=0, columnspan=2, row=6, pady=3)
        self._frame_configurations.grid(column=0, padx=10, pady=10, row=0, sticky="n")
        self._frame_colors = ttk.Frame(self._root, width=200)
        self._data_colors = ttk.Label(
            self._frame_colors, text="Cores dos dados dos visuais:"
        )
        self._data_colors.grid(column=0, columnspan=6, row=0)
        self._entry_color_1 = ttk.Entry(self._frame_colors, width=8)
        self._entry_color_1.grid(column=0, row=1, pady=3)
        self._entry_color_2 = ttk.Entry(self._frame_colors, width=8)
        self._entry_color_2.grid(column=1, row=1, pady=3)
        self._entry_color_3 = ttk.Entry(self._frame_colors, width=8)
        self._entry_color_3.grid(column=2, row=1, pady=3)
        self._entry_color_4 = ttk.Entry(self._frame_colors, width=8)
        self._entry_color_4.grid(column=3, row=1, pady=3)
        self._entry_color_5 = ttk.Entry(self._frame_colors, width=8)
        self._entry_color_5.grid(column=4, row=1, pady=3)
        self._entry_color_6 = ttk.Entry(self._frame_colors, width=8)
        self._entry_color_6.grid(column=5, row=1, pady=3)
        self._color_1 = ttk.Label(self._frame_colors, width=4)
        self._color_1.grid(column=0, row=2, pady=3)
        self._color_2 = ttk.Label(self._frame_colors, width=4)
        self._color_2.grid(column=1, row=2, pady=3)
        self._color_3 = ttk.Label(self._frame_colors, width=4)
        self._color_3.grid(column=2, row=2, pady=3)
        self._color_4 = ttk.Label(self._frame_colors, width=4)
        self._color_4.grid(column=3, row=2, pady=3)
        self._color_5 = ttk.Label(self._frame_colors, width=4)
        self._color_5.grid(column=4, row=2, pady=3)
        self._color_6 = ttk.Label(self._frame_colors, width=4)
        self._color_6.grid(column=5, row=2, pady=3)
        self._btn_change_colors = ttk.Button(
            self._frame_colors, command=self._change_colors, text="Alterar cores"
        )
        self._btn_change_colors.grid(column=0, columnspan=6, row=3, pady=3)
        self._frame_colors.grid(column=1, padx=10, pady=10, row=0, sticky="n")

    def _reset_values(self):
        self._entry_theme_name.delete(0, "end")
        self._combobox_font.set("")
        self._entry_background_color.delete(0, "end")
        self._entry_font_color.delete(0, "end")
        self._entry_font_size_values_card.delete(0, "end")
        self._entry_color_1.delete(0, "end")
        self._entry_color_2.delete(0, "end")
        self._entry_color_3.delete(0, "end")
        self._entry_color_4.delete(0, "end")
        self._entry_color_5.delete(0, "end")
        self._entry_color_6.delete(0, "end")
        self._color_1["background"] = ""
        self._color_2["background"] = ""
        self._color_3["background"] = ""
        self._color_4["background"] = ""
        self._color_5["background"] = ""
        self._color_6["background"] = ""

    def _load_theme(self):
        self._reset_values()

        try:
            with open(
                askopenfilename(
                    title="Selecinar tema", filetypes=[("Json files", ".json")]
                ),
                "r",
                encoding="utf-8-sig",
            ) as theme:
                self._data = json.loads(theme.read())
                self._define_default_values_entry()
        except FileNotFoundError:
            messagebox.showerror("Tema inválido", "Selecine um tema")

    def _define_default_values_entry(self):
        if self._data:
            self._entry_theme_name.insert(0, self._data["name"])
            self._entry_font_size_values_card.insert(
                0, self._data["textClasses"]["label"]["fontSize"]
            )
            self._combobox_font.set(self._data["textClasses"]["title"]["fontFace"])
            self._entry_font_color.insert(
                0, self._data["textClasses"]["title"]["color"]
            )

            if (
                "solid"
                in self._data["visualStyles"]["*"]["*"]["background"][0]["color"]
            ):
                self._entry_background_color.insert(
                    0,
                    self._data["visualStyles"]["*"]["*"]["background"][0]["color"][
                        "solid"
                    ]["color"],
                )
            else:
                self._entry_background_color.insert(
                    0, self._data["visualStyles"]["*"]["*"]["background"][0]["color"]
                )

            self._entry_color_1.insert(0, self._data["dataColors"][0])
            self._entry_color_2.insert(0, self._data["dataColors"][1])
            self._entry_color_3.insert(0, self._data["dataColors"][2])
            self._entry_color_4.insert(0, self._data["dataColors"][3])
            self._entry_color_5.insert(0, self._data["dataColors"][4])
            self._entry_color_6.insert(0, self._data["dataColors"][5])

            self._color_1["background"] = self._data["dataColors"][0]
            self._color_2["background"] = self._data["dataColors"][1]
            self._color_3["background"] = self._data["dataColors"][2]
            self._color_4["background"] = self._data["dataColors"][3]
            self._color_5["background"] = self._data["dataColors"][4]
            self._color_6["background"] = self._data["dataColors"][5]

    def _is_valid_color(self, color, type_):
        color = color.replace("\t", "").replace("\n", "").replace(" ", "").strip()
        if type_ == "config":
            default_color_error = "#000000"
            if re.search(r"^#(?:[0-9a-fA-F]{3}){1,2}$", color):
                self.show_error = False
                return color

            if not self.show_error:
                messagebox.showerror("Cor inválida", "Insira uma cor válida")
                self.show_error = True

            return default_color_error
        else:
            if re.search(r"^#(?:[0-9a-fA-F]{3}){1,2}$", color):
                return color
            messagebox.showerror("Cor inválida", "Insira uma cor válida")

    def _is_valid_font(self, font):
        if font in self._power_bi_fonts:
            return True
        messagebox.showerror("Fonte inválida", "Selecione uma fonte válida")

    def _is_valid_size_font(self, number):
        try:
            number = float(number)
            if 45 >= number >= 11:
                return float(number)
            else:
                messagebox.showerror(
                    "Tamanho não permitido", "Insira um tamanho entre 11 e 45"
                )
        except ValueError:
            messagebox.showerror(
                "Tamanho inválido", "Insira um tamanho válido para a fonte"
            )

    def _validate_entrys(self):
        if (
            self._entry_background_color.get()
            and self._entry_theme_name.get()
            and self._entry_font_color.get()
            and self._entry_font_size_values_card.get()
        ):
            if (
                self._is_valid_color(
                    str(self._entry_background_color.get()), "generate"
                )
                and self._is_valid_color(str(self._entry_font_color.get()), "generate")
                and self._is_valid_size_font(self._entry_font_size_values_card.get())
                and self._is_valid_font(self._combobox_font.get())
            ):
                self._data["name"] = self._entry_theme_name.get()
                if (
                    "solid"
                    in self._data["visualStyles"]["*"]["*"]["background"][0]["color"]
                ):
                    self._data["visualStyles"]["*"]["*"]["background"][0]["color"][
                        "solid"
                    ]["color"] = self._rm_spaces_and_special_characters_colors(
                        self._entry_background_color.get()
                    )
                else:
                    self._data["visualStyles"]["*"]["*"]["background"][0]["color"] = (
                        self._rm_spaces_and_special_characters_colors(
                            self._entry_background_color.get()
                        )
                    )
                self._data["textClasses"]["title"]["color"] = (
                    self._rm_spaces_and_special_characters_colors(
                        self._entry_font_color.get()
                    )
                )
                self._data["textClasses"]["callout"]["fontSize"] = int(
                    self._entry_font_size_values_card.get()
                )
                self._data["textClasses"]["title"][
                    "fontFace"
                ] = self._combobox_font.get()
                return True
        else:
            messagebox.showerror("Campos vazios", "Preencha todos os campos")
            return False

    def _rm_spaces_and_special_characters_colors(self, color):
        return color.replace("\t", "").replace("\n", "").replace(" ", "").strip()

    def _change_colors(self):
        self.show_error = False

        self._data["dataColors"][0] = self._is_valid_color(
            self._entry_color_1.get(), "config"
        )

        self._color_1["background"] = self._rm_spaces_and_special_characters_colors(
            self._data["dataColors"][0]
        )

        self._data["dataColors"][1] = self._is_valid_color(
            self._entry_color_2.get(), "config"
        )

        self._color_2["background"] = self._rm_spaces_and_special_characters_colors(
            self._data["dataColors"][1]
        )

        self._data["dataColors"][2] = self._is_valid_color(
            self._entry_color_3.get(), "config"
        )

        self._color_3["background"] = self._rm_spaces_and_special_characters_colors(
            self._data["dataColors"][2]
        )

        self._data["dataColors"][3] = self._is_valid_color(
            self._entry_color_4.get(), "config"
        )

        self._color_4["background"] = self._rm_spaces_and_special_characters_colors(
            self._data["dataColors"][3]
        )

        self._data["dataColors"][4] = self._is_valid_color(
            self._entry_color_5.get(), "config"
        )

        self._color_5["background"] = self._rm_spaces_and_special_characters_colors(
            self._data["dataColors"][4]
        )

        self._data["dataColors"][5] = self._is_valid_color(
            self._entry_color_6.get(), "config"
        )

        self._color_6["background"] = self._rm_spaces_and_special_characters_colors(
            self._data["dataColors"][5]
        )

    def _generate_theme(self):
        ok = self._validate_entrys()

        if ok:
            path = asksaveasfilename(
                title="Salvar como",
                filetypes=[("Json files", ".json")],
                defaultextension=".json",
                initialfile=f"{self._data["name"]}.json",
            )

            if path:
                root, ext = os.path.splitext(path)

                if root.endswith(".json"):
                    path = root

                with open(f"{path}", mode="w", encoding="utf-8") as result:
                    json.dump(self._data, result)
                    messagebox.showinfo("Arquivo salvo", "Arquivo salvo com sucesso")

    def _center_window(self):
        screen_width = self._root.winfo_screenwidth()
        screen_height = self._root.winfo_screenheight()
        x = (screen_width / 2) - (self._width / 2)
        y = (screen_height / 2) - (self._height / 2)
        self._root.geometry(f"{self._width}x{self._height}+{int(x)}+{int(y)}")

    def run(self):
        self._root.mainloop()


if __name__ == "__main__":
    app = ThemesGenerator()
    app.run()
