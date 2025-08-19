"""
Plotting Module for FTIR Spectroscopy

Contains functions for setting up OriginLab-style plotting and
formatting scientific plots for publication quality.
"""

import matplotlib.pyplot as plt
from cycler import cycler


def setup_originlab_style():
    """Set up OriginLab-like plotting style"""
    # Set figure background to white
    plt.rcParams["figure.facecolor"] = "white"
    plt.rcParams["axes.facecolor"] = "white"

    # Set font properties (OriginLab typically uses Arial/Helvetica)
    plt.rcParams["font.family"] = ["Arial", "DejaVu Sans", "sans-serif"]
    plt.rcParams["font.size"] = 12
    plt.rcParams["axes.titlesize"] = 14
    plt.rcParams["axes.labelsize"] = 12
    plt.rcParams["xtick.labelsize"] = 11
    plt.rcParams["ytick.labelsize"] = 11
    plt.rcParams["legend.fontsize"] = 10

    # Set line and marker properties
    plt.rcParams["lines.linewidth"] = 1.5
    plt.rcParams["axes.linewidth"] = 1.0

    # Set tick properties (OriginLab style)
    plt.rcParams["xtick.direction"] = "in"
    plt.rcParams["ytick.direction"] = "in"
    plt.rcParams["xtick.major.size"] = 4
    plt.rcParams["ytick.major.size"] = 4
    plt.rcParams["xtick.minor.size"] = 2
    plt.rcParams["ytick.minor.size"] = 2
    plt.rcParams["xtick.major.width"] = 1.0
    plt.rcParams["ytick.major.width"] = 1.0

    # Set grid properties
    plt.rcParams["grid.color"] = "lightgray"
    plt.rcParams["grid.linestyle"] = "-"
    plt.rcParams["grid.linewidth"] = 0.5
    plt.rcParams["grid.alpha"] = 0.7

    # Set spine properties
    plt.rcParams["axes.spines.left"] = True
    plt.rcParams["axes.spines.bottom"] = True
    plt.rcParams["axes.spines.top"] = True
    plt.rcParams["axes.spines.right"] = True

    # Color cycle similar to OriginLab
    colors = [
        "#000000",  # Black
        "#FF0000",  # Red
        "#0000FF",  # Blue
        "#008000",  # Green
        "#FF8000",  # Orange
        "#800080",  # Purple
        "#008080",  # Teal
        "#808000",  # Olive
        "#FF00FF",  # Magenta
        "#800000",  # Maroon
    ]
    plt.rcParams["axes.prop_cycle"] = cycler("color", colors)


def format_originlab_plot(ax, title, xlabel, ylabel, show_minor_ticks=True):
    """
    Format a plot to look like OriginLab
    
    Parameters:
    ax: matplotlib axis object
    title: plot title
    xlabel: x-axis label
    ylabel: y-axis label
    show_minor_ticks: whether to show minor tick marks
    """
    # Set title and labels
    ax.set_title(title, fontweight="bold", pad=20)
    ax.set_xlabel(xlabel, fontweight="bold")
    ax.set_ylabel(ylabel, fontweight="bold")

    # Set tick parameters
    ax.tick_params(axis="both", which="major", direction="in", length=4, width=1)
    ax.tick_params(axis="both", which="minor", direction="in", length=2, width=1)

    # Show minor ticks
    if show_minor_ticks:
        ax.minorticks_on()

    # Set spine properties
    for spine in ax.spines.values():
        spine.set_linewidth(1.0)
        spine.set_color("black")

    # Grid
    ax.grid(True, which="major", alpha=0.3, linewidth=0.5)
    ax.grid(True, which="minor", alpha=0.1, linewidth=0.3)

    # Legend with OriginLab style
    if ax.get_legend():
        legend = ax.legend(
            frameon=True,
            fancybox=False,
            shadow=False,
            framealpha=1.0,
            edgecolor="black",
            facecolor="white",
        )
        legend.get_frame().set_linewidth(1.0)


def create_originlab_legend(ax):
    """
    Create an OriginLab-style legend for the plot
    
    Parameters:
    ax: matplotlib axis object
    
    Returns:
    legend: matplotlib legend object
    """
    legend = ax.legend(
        frameon=True,
        fancybox=False,
        shadow=False,
        framealpha=1.0,
        edgecolor="black",
        facecolor="white",
    )
    legend.get_frame().set_linewidth(1.0)
    return legend