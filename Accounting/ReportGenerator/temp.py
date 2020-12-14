def file_save():
    f = filedialog.asksaveasfile(mode='w', defaultextension=".txt")
    if f is None:  # asksaveasfile return `None` if dialog closed with "cancel".
        return
    text2save = str(f.text.get(1.0, f.END))  # starts from `1.0`, not `0.0`
    f.write(text2save)
    f.close()  # `()` was missing.
