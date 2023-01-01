import tkinter as tk
from ExcelVlookupForTranslation import ExcelVlookupForTranslation


class MainWindow:
    def __init__(self, window):
        self.window = window
        self.window.title('Main Window')
        self.widgets = []
        self.create_widgets()

    def create_widgets(self):
        window_width = 350
        window_height = 250
        screen_width = self.window.winfo_screenwidth()
        screen_height = self.window.winfo_screenheight()
        center_x = int(screen_width / 2 - window_width / 2)
        center_y = int(screen_height / 2 - window_height / 2)
        self.window.geometry(f'{window_width}x{window_height}+{center_x}+{center_y}')

        label = tk.Label(text='Choose the program:')
        label.pack(pady=10)
        self.widgets.append(label)

        button = tk.Button(text="Ready Excel File For Translation",
                           command=self.switch_to_excel_vlookup_for_translation)
        button.pack(pady=10)
        self.widgets.append(button)

    def clear_window(self):
        for widget in self.widgets:
            widget.destroy()
        self.widgets = []

    def switch_to_excel_vlookup_for_translation(self):
        self.clear_window()
        excel_window = ExcelVlookupForTranslation(window)


if __name__ == '__main__':
    window = tk.Tk()
    main_window = MainWindow(window)
    window.mainloop()
