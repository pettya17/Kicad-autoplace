import os

import pcbnew
import wx

from .example_dialog import ExampleDialog


class TemplatePluginAction(pcbnew.ActionPlugin):
    def defaults(self) -> None:
        self.name = "PCB from Xlsx plugin"
        self.category = "Utilitie"
        self.description = "This plugin include parts to the opened PCB from excel table."
        self.show_toolbar_button = True
        self.icon_file_name = os.path.join(os.path.dirname(__file__), "icon.png")

    def Run(self) -> None:
        pcb_frame = next(
            x for x in wx.GetTopLevelWindows() if x.GetName() == "PcbFrame"
        )

        dlg = ExampleDialog(pcb_frame)
        dlg.ShowModal()
        pass
        dlg.Destroy()
