# color
def generate_color_codes(start_color, end_color, num_colors):
    """
    Generate a list of color codes between start_color and end_color.

    Arguments:
    start_color -- a hex color code string for the starting color, e.g. "#FF0000" for red
    end_color -- a hex color code string for the ending color
    num_colors -- the number of colors to generate in the list

    Returns:
    A list of color codes in the format "#RRGGBB"
    """
    # Convert the hex color codes to RGB tuples
    start_r, start_g, start_b = tuple(int(start_color[i:i+2], 16) for i in (1, 3, 5))
    end_r, end_g, end_b = tuple(int(end_color[i:i+2], 16) for i in (1, 3, 5))

    color_codes = []
    for i in range(num_colors):
        # Calculate the RGB values for the current color
        r = start_r + (i * (end_r - start_r)) // (num_colors - 1)
        g = start_g + (i * (end_g - start_g)) // (num_colors - 1)
        b = start_b + (i * (end_b - start_b)) // (num_colors - 1)

        # Convert the RGB values to a hex string
        color_code = "#{:02x}{:02x}{:02x}".format(r, g, b)

        # Add the color code to the list
        color_codes.append(color_code)

    return color_codes
