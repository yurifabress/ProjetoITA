import tkinter as tk


class LoginApp:

  def __init__(self, root, open_inventory_app_callback):
    self.root = root
    self.root.title("ITAMIL - TI")
    self.root.geometry("600x250")

    self.label = tk.Label(root,
                          text="Itamil Control System",
                          font=("Helvetica", 18))
    self.label.pack(pady=20)

    self.username_label = tk.Label(root,
                                   text="Usuário:",
                                   font=("Helvetica", 12))
    self.username_label.pack()
    self.username_entry = tk.Entry(root, font=("Helvetica", 12))
    self.username_entry.pack(pady=5)

    self.password_label = tk.Label(root, text="Senha:", font=("Helvetica", 12))
    self.password_label.pack()
    self.password_entry = tk.Entry(root, show="*")
    self.password_entry.pack(pady=5)

    self.login_button = tk.Button(root,
                                  text="Login",
                                  command=self.login,
                                  font=("Helvetica", 12),
                                  bg="#4CAF50",
                                  fg="white")
    self.login_button.pack(pady=20)

    self.open_inventory_app_callback = open_inventory_app_callback

  def login(self):
    username = self.username_entry.get()
    password = self.password_entry.get()

    if username == "admin" and password == "admin123":
      self.root.destroy()
      self.open_inventory_app_callback()
    else:
      self.error_label.config(text="Usuário ou senha incorretos",
                              foreground="red")
