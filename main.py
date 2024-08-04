import tkinter as tk
from interface import VoxBuddyInterface
from commands import process_command_func

if __name__ == "__main__":
    root = tk.Tk()
    interface = VoxBuddyInterface(root, process_command_func)
    root.mainloop()
