import tkinter as tk
from LoginApp import LoginApp
from TIInventoryApp import TIInventoryApp


def open_inventory_app():
  root = tk.Tk()
  app = TIInventoryApp(root)
  root.mainloop()


if __name__ == "__main__":
  root = tk.Tk()
  login_app = LoginApp(root, open_inventory_app)
  root.mainloop()
