import tkinter as tk
from tkinter.font import Font
import webbrowser
from pathlib import Path  # to convert file path to url

__version__ = '0.1'

from svg_from_pdf import export_active_chart_to_svg

class AppWindow(tk.Frame):


    def _open_in_webrowser(self, url):
        webbrowser.open_new(url)

    def __init__(self):

        self.master = tk.Tk()
        tk.Frame.__init__(self, self.master)

        self.pack(fill=tk.BOTH, expand=1)

        text = tk.Label(self, text="Select chart in active sheet in Excel then press button")
        text.place(x=10, y=10)
        self._Font = text['font'] # pick up the default font

        svg_btn = tk.Button(self,
                           text="Create SVG",
                           command=self.svg)

        svg_btn.place(y=60, relx=0.5, anchor= tk.CENTER)

        self.github_link = tk.Label(self, text="GitHub", fg="blue", cursor="hand2")
        self.github_link.bind("<Button-1>", lambda e: self._open_in_webrowser("https://github.com/jmills147/excel_to_svg"))
        self.github_link.place(x=250, y=90)

        self.version_lbl = tk.Label(self, text='v' + __version__)
        self.version_lbl.place(x=265, y=70)

        self.master.wm_title("Excel to SVG")
        self.master.geometry("300x120")
        self.master.resizable(False, False)
        self.master.iconbitmap('ladybird.ico')
        self.mainloop()


    def post_svg_winow_update(self, sz_svg_path):


        # SVG saved to label
        svg_saved_to_lbl = tk.Label(self,
                                wraplength=180,
                                text=r"SVG saved to:")

        svg_saved_to_lbl.place(x=10, y=90)


        # SVG path text box
        w = tk.Text(self,
                    height=5,
                    width=45,
                    background=self['bg'],
                    font = self._Font,
                    borderwidth=2,
                    relief="sunken")

        w.place(x=10, y=120)

        w.configure(state="normal")
        w.delete(1.0, "end")
        w.insert(1.0, sz_svg_path)
        w.configure(state="disabled")


        # Open in browser button
        svg_url = Path(sz_svg_path).as_uri()

        open_svg_btn = tk.Button(self,
                           text="Open SVG in Webrowser",
                           command=lambda: self._open_in_webrowser(str(svg_url)))

        open_svg_btn.place(x=10, y=220)


        # Move github link
        self.github_link.place(x=250, y=210)

        # Move version label
        self.version_lbl.place(x=265, y=230)

        # Resize window
        self.master.geometry("300x260")  # 230

    def svg(self):

        sz_svg_path = export_active_chart_to_svg()

        self.post_svg_winow_update(sz_svg_path)


if __name__ == '__main__':
    AppWindow()