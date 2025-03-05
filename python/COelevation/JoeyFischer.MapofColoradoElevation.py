"""
File: Fischer_Project9.py.py
Author: Joey Fischer
Date: November 12, 2024
Description: This program takes CO elevation data and draws a window with each elevation as a pixel where the intensity is based on 
the elevation of that point (white being maximum elevation and black being minimum elevation). Clicking a point tells the elevation
at that point and pressing any key closes the window.
"""

import dudraw


# function to find the max elevation by using a nested loop to iterate over the 2D list
def max_elev(data):
    max = 0
    # each element of outer list
    for i in range(len(data)):
        #each element of inner lists
        for j in range(len(data[i])):
            # check if value is bigger to set as new maximum
            if data[i][j] > max:
                max = data[i][j]
    #return the maximum
    return max

# function to find the min elevation by using a nested loop to iterate over the 2D list
def min_elev(data):
    min = float('inf')
    # each element of outer list
    for i in range(len(data)):
        #each element of inner lists
        for j in range(len(data[i])):
            # check if value is smaller to set as new minimum
            if data[i][j] < min:
                min = data[i][j]
    #return the minimum
    return min

def draw_elevation(data,max_val,min_val):
    for row in range(len(data)):
        for col in range(len(data[row])):
            elevation = data[row][col]
            gradient = int(((max_val - elevation) * 255)/(max_val-min_val))
            dudraw.set_pen_color_rgb(min(gradient,255),min(gradient,255),min(gradient,255))
            dudraw.point(col+1,row+1)

def main():    
    #empty list for elevation data
    elevation_data = []
    # open file and create 2D list out of values
    with open('CO_elevations_feet.txt', 'r') as file:
        for line in file:
            row = list(map(int, line.split()))
            elevation_data.append(row)
    # finds maximum elevation
    max_val = max_elev(elevation_data)
    # finds minimum elevation
    min_val = min_elev(elevation_data)
    # set canvas size and scale
    dudraw.set_canvas_size(560,560)
    dudraw.set_x_scale(0,560)
    dudraw.set_y_scale(0,560)
    dudraw.clear()
    dudraw.show()
    # initiate last clicked elevation variable
    last_clicked_elevation = None
    while True:
        # check if mouse clicked
        if dudraw.mouse_clicked():
            # set position of mouse click
            mouse_x = int(dudraw.mouse_x())
            mouse_y = int(dudraw.mouse_y())
            # set last clicked elevation, making sure to keep in the bounds
            if 0 <= mouse_x < 560 and 0 <= mouse_y < 560:
                last_clicked_elevation = elevation_data[mouse_y][mouse_x]
        # clear and draw elevation
        dudraw.clear()
        draw_elevation(elevation_data, max_val, min_val)
        # print elevation in bottom right corner in a box
        if last_clicked_elevation is not None:
            dudraw.set_pen_color(dudraw.WHITE)
            dudraw.filled_rectangle(560-50,25,50,25)
            dudraw.set_pen_color(dudraw.BLACK)
            dudraw.text(560 -50, 25, f"Elevation: {last_clicked_elevation}")
            dudraw.rectangle(560 - 50, 25,50,25)
        # quit the window if any key is pressed
        if dudraw.has_next_key_typed():  
            break
        # show, updating frequently
        dudraw.show(0.1) 

#run function
if __name__ == '__main__':
    main()