
import tkinter as tk

root = tk.Tk()
t = tk.Text(root)
t.pack()

# Chèn 3 dòng text
t.insert("end", "A\nBOLD HERE\nC\n")

# Gán tag 'bold' cho chữ 'BOLD HERE' (nằm ở dòng 2)
t.tag_add("bold", "2.0", "2.9")  # 9 ký tự: "BOLD HERE"

# Gán thêm tag 'bold' cho chữ 'C' (dòng 3, cột 0 -> cột 1)
t.tag_add("bold", "3.0", "3.1")

# Lấy ranges của tag 'bold'
ranges = t.tag_ranges("bold")
print(ranges)  # sẽ in một list các tk index objects, nhưng bạn dùng như chuỗi được

# Hiển thị từng cặp start-end và nội dung tương ứng
for i in range(0, len(ranges), 2):
    start = ranges[i]
    end = ranges[i+1]
    content = t.get(start, end)
    print(f"Range {i//2+1}: {start} -> {end} | '{content}'")

root.mainloop()
