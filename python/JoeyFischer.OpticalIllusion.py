import dudraw
# Hermann Scintillating Grid Illusion with different colors

def main():
    #set canvas and background
    dudraw.set_canvas_size(500,500)
    dudraw.clear_rgb(0,100,0)
    x = 0
    #draw horizontal lines at each point
    for i in range(9):
        dudraw.set_pen_color(dudraw.MAGENTA)
        dudraw.line(x,0,x,1)
        x += 1/8
        x2 = x
        x3 = x
        # draw a bunch of lines around each line to thicken them
        for j in range(8):
            dudraw.line(x2,0,x2,1)
            x2 -= .001
        # draw a bunch of lines around each line to thicken them
        for k in range(8):
            dudraw.line(x3,0,x3,1)
            x3 += .001
    y = 0
    #draw vertical lines at each point
    for i in range(9):
        dudraw.set_pen_color_rgb(255,0,255)
        dudraw.line(0,y,1,y)
        y += 1/8
        y2 = y
        y3 = y
        # draw a bunch of lines around each line to thicken them
        for j in range(8):
            dudraw.line(0,y2,1,y2)
            y2 -= .001
        # draw a bunch of lines around each line to thicken them
        for k in range(8):
            dudraw.line(0,y3,1,y3)
            y3 += .001
    y_loc = 1/8
    # draw circles at each intersection
    # nested loops to draw across all intersectiosn 
    for i in range(7):
        x_loc = 1/8
        for i in range(7):
            dudraw.set_pen_color_rgb(255,153,255)
            dudraw.filled_circle(x_loc,y_loc,.01)
            x_loc += 1/8
        y_loc += 1/8
    # show drawing
    dudraw.show(float('inf'))

if __name__ == "__main__":
    main()