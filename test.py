from tkinter import Tk, Canvas
from time import sleep
from random import randint as rand
# num_squares = int(input("How many squares?"))
window = Tk()
window.title("The window")
canvas = Canvas(window,width=400,height=400, bg="black")
num_squares = int(input("How many squares?"))
square = []
colour = ["red","yellow","green","blue"]

for i in range(num_squares):
    c_col = rand(0,3)
    x = rand(10, 300)
    y = rand(10, 300)
    xy = (x, y, x+10, y+10)

    square.append(canvas.create_rectangle(xy, fill=colour[c_col]))
canvas.pack()

x = [1] * num_squares
y = [1] * num_squares

while True:
    for i in range(num_squares):
        pos = canvas.coords(square[i])
        if pos[3] > 400 or pos[1] < 0:
            y[i] = -y[i]
        if pos[0] < 0 or pos[2] > 400:
            x[i] = -x[i]

        for j in range(num_squares):
            if i == j: continue
            pos2 = canvas.coords(square[j])
            if pos[0] < pos2[2] and pos[2] >pos2[0] and pos[1] < pos2[3] and pos[3] > pos2[1]:
                y[i] = -y[i]
                x[i] = -x[i]
                y[j] = -y[j]
                x[j] = -x[j]
        canvas.move(square[i], x[i], y[i])
    sleep(0.002)
    window.update()
window.mainloop()