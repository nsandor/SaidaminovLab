import csv
import numpy as np
from scipy import signal
import matplotlib.pyplot as plt
from scipy.sparse import csc_matrix
from scipy.sparse import spdiags
import scipy.sparse.linalg as spla
import pandas as pd


# Remove first {time_offset} seconds
def offset_XY(X,Y,time_offset):    
    x = [n-time_offset for n in X]
    k = 0
    for i in range(len(x)):
        if x[i] < 0:
            k += 1
    x = x[k:]
    y = Y[k:]
    return x, y

# Baseline estimation by AsLS
def baseline_als(y, lam, p, niter=10):
    #https://stackoverflow.com/questions/29156532/python-baseline-correction-library
    #p: 0.001 - 0.1, lam: 10^2 - 10^9
    # Baseline correction with asymmetric least squares smoothing, P. Eilers, 2005
    L = len(y)
    D = csc_matrix(np.diff(np.eye(L), 2))
    w = np.ones(L)
    for i in range(niter):
        W = spdiags(w, 0, L, L)
        Z = W + lam * D.dot(D.transpose())
        z = spla.spsolve(Z, w*y)
        w = p * (y > z) + (1-p) * (y < z)
    return z

# Check the baseline
def baseline_test(X,Y,paramAsLS):
    # baseline estimation and smoothing
    Y_np = np.array(Y)
    bkg = baseline_als(Y_np,paramAsLS[0], paramAsLS[1])
    fix = Y_np - bkg

    #figures
    plt.figure(figsize=(8,6))
    ax1 = plt.subplot2grid((2,2), (0,0), colspan=2)
    ax2 = plt.subplot2grid((2,2), (1,0), colspan=2)

    ax1.plot(X, Y, linewidth=2)
    ax1.plot(X, bkg, "b", linewidth=1, linestyle = "dashed")

    ax2.plot(X, fix, "g", linewidth=1, linestyle = "-")

    plt.axis("tight")
    plt.xlabel('Time (s)')
    plt.show()
    return bkg, fix
    
# Output csv file and figure
def outFigCSV(X,Y,paramAsLS, save_path, save_option):
    # # save path
    save_path_csv = f'{save_path}.csv'
    save_path_fig = f'{save_path}.jpg'
    # baseline estimation and smoothing
    Y_np = np.array(Y)
    bkg = baseline_als(Y_np,paramAsLS[0], paramAsLS[1])
    fix = Y_np - bkg
    
    # output
    if save_option:
        save_data = pd.DataFrame(
            data = {
                "Time (s)": X,
                "J (nA/cm2)": Y,
                "bkg": bkg,
                "J-bkg": fix
            }
        )
        save_data.to_csv(save_path_csv)

    #figures
    plt.figure(figsize=(12,9))
    ax1 = plt.subplot2grid((2,2), (0,0), colspan=2)
    ax2 = plt.subplot2grid((2,2), (1,0), colspan=2)

    ax1.plot(X, Y, linewidth=2)
    ax1.plot(X, bkg, "b", linewidth=1, linestyle = "dashed", label="baseline")

    ax2.plot(X, fix, "g", linewidth=1, linestyle = "-", label="remove baseline")

    plt.axis("tight")
    plt.legend(loc='best', frameon = False)
    plt.xlabel('Time (s)')
    if save_option:
        plt.savefig(save_path_fig, dpi=1200, bbox_inches='tight')
    plt.show()
    
def ItoJ(I, area, unit):
    '''
    I is a list of the current (A)
    unit options: 'pA/cm2', 'nA/cm2', 'µA/cm2', 'mA/cm2'
    area should be cm2
    '''
    if unit == 'pA/cm2':
        J = [abs(n/area*1e12) for n in I]
        J_unit = "Current density (pA $\mathrm{cm^{-2}}$)"
    elif unit == 'nA/cm2':
        J = [abs(n/area*1e9) for n in I]
        J_unit = "Current density (nA $\mathrm{cm^{-2}}$)"
    elif unit == 'µA/cm2':
        J = [abs(n/area*1e6) for n in I]
        J_unit = "Current density (µA $\mathrm{cm^{-2}}$)"
    elif unit == 'mA/cm2':
        J = [abs(n/area*1e3) for n in I]
        J_unit = "Current density (mA $\mathrm{cm^{-2}}$)"
    return J, J_unit

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