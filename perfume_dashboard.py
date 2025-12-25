import pandas as pd
import tkinter as tk
from tkinter import ttk
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import matplotlib.pyplot as plt
import seaborn as sns

# Load and preprocess data with error handling
try:
    df = pd.read_excel("perfumedata.xlsx")
except FileNotFoundError:
    print("Error: The file 'perfumedata.xlsx' was not found.")
    exit(1)

# Check required columns
required_columns = ['Causes/Effects', 'Flavour of Aroma', 'Food', 'Synthetic Aroma Agent Name', 'Natural Agents']
for col in required_columns:
    if col not in df.columns:
        print(f"Error: Required column '{col}' is missing from the DataFrame.")
        exit(1)

# Preprocessing the 'Causes/Effects' column
replacement_mapping = {
    'irritation effects': 'Irritation Effects',
    'specialized effects': 'Specialized Effects',
    'specified effects': 'Specialized Effects',  # Assuming this is a typo
    'systemic toxicity': 'Systemic Toxicity',
    'safety approvals': 'Safety Approvals',
    'allergic reaction': 'Allergic Reaction'
}

# Normalize and replace values
df['Causes/Effects'] = df['Causes/Effects'].str.lower()
df['Causes/Effects'] = df['Causes/Effects'].replace(replacement_mapping)

# Ensure the column names are stripped of whitespace
df.columns = df.columns.str.strip()

# Function to analyze systemic toxicity
def analyze_systemic_toxicity(df):
    # Filter data for systemic toxicity
    toxicity_data = df[df['Causes/Effects'].str.contains('Systemic Toxicity', case=False)]

    # Count the number of synthetic aroma agents causing systemic toxicity
    synthetic_count = len(toxicity_data)
    print(f'Number of Synthetic Aroma Agents Causing Systemic Toxicity: {synthetic_count}')

    # Get unique flavours and natural agents
    flavours = toxicity_data['Flavour of Aroma'].unique()
    natural_agents = toxicity_data['Natural Agents'].unique()

    print("\nFlavours of Synthetic Aroma Agents Causing Systemic Toxicity:")
    print(flavours)

    print("\nNatural replacements of Synthetic Aroma Agents Causing Systemic Toxicity:")
    replacement_table = toxicity_data[['Synthetic Aroma Agent Name', 'Natural Agents']]
    print(replacement_table)

    # Count occurrences of food items
    flavour_counts = toxicity_data['Food'].value_counts()

    # Create a pie chart for visual representation
    plt.figure(figsize=(8, 8))
    flavour_counts.plot(kind='pie', autopct='%1.1f%%', startangle=90, colors=plt.cm.Paired.colors)
    plt.title('Foods using Synthetic Aroma Agents Causing Systemic Toxicity')
    plt.ylabel('')

    # Return the figure for embedding in the GUI
    return plt.gcf()

# Function to analyze safety approvals
def analyze_safety_approvals(df):
    # Filter data for safety approvals
    safety_approval_data = df[df['Causes/Effects'].str.contains('Safety Approval', case=False)]

    # Count the number of synthetic aroma agents with safety approvals
    synthetic_count = len(safety_approval_data)
    print(f'Number of Synthetic Aroma Agents with Safety Approvals: {synthetic_count}')

    # Get unique flavours and natural agents
    flavours = safety_approval_data['Flavour of Aroma'].unique()
    natural_agents = safety_approval_data['Natural Agents'].unique()

    print("\nFlavours of Synthetic Aroma Agents with Safety Approvals:")
    print(flavours)

    print("\nNatural replacements of Synthetic Aroma Agents with Safety Approvals:")
    replacement_table = safety_approval_data[['Synthetic Aroma Agent Name', 'Natural Agents']]
    print(replacement_table)

    # Count occurrences of food items
    flavour_counts = safety_approval_data['Food'].value_counts()

    # Create a pie chart for visual representation
    plt.figure(figsize=(8, 8))
    flavour_counts.plot(kind='pie', autopct='%1.1f%%', startangle=90, colors=plt.cm.Paired.colors)
    plt.title('Foods using Synthetic Aroma Agents with Safety Approvals')
    plt.ylabel('')

    # Return the figure for embedding in the GUI
    return plt.gcf()

# New function to plot distribution of effects
def plot_effect_distribution():
    plt.clf()
    effect_counts = df['Causes/Effects'].value_counts()
    plt.figure(figsize=(10, 6))
    effect_counts.plot(kind='bar')
    plt.title('Distribution of Effects of Synthetic Aromatic Agents')
    plt.xlabel('Effects')
    plt.ylabel('Count')
    plt.xticks(rotation=45)
    return plt.gcf()

# New function to plot heatmap of effects by flavor
def plot_flavor_effect_heatmap():
    plt.clf()
    flavor_effect_counts = df.groupby(['Flavour of Aroma', 'Causes/Effects']).size().reset_index(name='Count')
    pivot_table = flavor_effect_counts.pivot(index='Flavour of Aroma', columns='Causes/Effects', values='Count').fillna(0)

    plt.figure(figsize=(12, 8))
    sns.heatmap(pivot_table, annot=True, fmt=".0f", cmap="YlGnBu")
    plt.title('Effects by Flavour of Aroma')
    plt.xlabel('Causes/Effects')
    plt.ylabel('Flavour of Aroma')
    plt.xticks(rotation=45)
    plt.tight_layout()
    return plt.gcf()

# New function to plot heatmap of effects by food type and correlation
def plot_food_effect_heatmap():
    plt.clf()
    food_effect_counts = df.groupby(['Food', 'Causes/Effects']).size().reset_index(name='Count')
    pivot_table = food_effect_counts.pivot(index='Food', columns='Causes/Effects', values='Count').fillna(0)

    # First heatmap for effects by food type
    plt.figure(figsize=(16, 10))
    sns.heatmap(pivot_table, annot=True, fmt=".0f", cmap="YlGnBu", cbar_kws={'label': 'Count'})
    plt.title('Effects by Food Type')
    plt.xlabel('Effects')
    plt.ylabel('Food Type')
    plt.xticks(rotation=45, ha='right')
    plt.tight_layout()
    food_effect_fig = plt.gcf()  # Store the figure for embedding

    # Calculate the correlation matrix
    correlation_matrix = pivot_table.corr()

    # Second heatmap for correlation
    plt.figure(figsize=(10, 8))
    sns.heatmap(correlation_matrix, annot=True, fmt=".2f", cmap="coolwarm", cbar_kws={'label': 'Correlation'})
    plt.title('Correlation Between Food Types and Effects')
    plt.tight_layout()
    correlation_fig = plt.gcf()  # Store the figure for embedding

    return food_effect_fig, correlation_fig  # Return both figures

# New function to analyze irritation effects and plot data
def plot_irritation_effects():
    plt.clf()
    irritation_data = df[df['Causes/Effects'].str.contains('Irritation Effects', case=False)]

    # Count the number of synthetic aroma agents causing irritation
    synthetic_count = len(irritation_data)
    print(f'Number of Synthetic Aroma Agents Causing Irritation: {synthetic_count}')

    # Get unique flavours and natural agents
    flavours = irritation_data['Flavour of Aroma'].unique()
    natural_agents = irritation_data['Natural Agents'].unique()

    print("\nFlavours of Synthetic Aroma Agents Causing Irritation:")
    print(flavours)

    print("\nNatural replacements of Synthetic Aroma Agents Causing Irritation:")
    replacement_table = irritation_data[['Synthetic Aroma Agent Name', 'Natural Agents']]
    print(replacement_table)

    # Count occurrences of food items
    flavour_counts = irritation_data['Food'].value_counts()

    # Create a pie chart for visual representation
    plt.figure(figsize=(8, 8))
    flavour_counts.plot(kind='pie', autopct='%1.1f%%', startangle=90, colors=plt.cm.Paired.colors)
    plt.title('Foods using Synthetic Aroma Agents Causing Irritation')
    plt.ylabel('')
    plt.tight_layout()
    return plt.gcf()  # Return the figure for embedding

# New function to analyze specialized effects and plot data
def plot_specialized_effects():
    plt.clf()
    specialized_data = df[df['Causes/Effects'].str.contains('Specialized Effects', case=False)]

    # Count the number of synthetic aroma agents causing specialized effects
    synthetic_count = len(specialized_data)
    print(f'Number of Synthetic Aroma Agents Causing Specialized Effects: {synthetic_count}')

    # Get unique flavours and natural agents
    flavours = specialized_data['Flavour of Aroma'].unique()
    natural_agents = specialized_data['Natural Agents'].unique()

    print("\nFlavours of Synthetic Aroma Agents Causing Specialized Effects:")
    print(flavours)

    print("\nNatural replacements of Synthetic Aroma Agents Causing Specialized Effects:")
    replacement_table = specialized_data[['Synthetic Aroma Agent Name', 'Natural Agents']]
    print(replacement_table)

    # Count occurrences of food items
    flavour_counts = specialized_data['Food'].value_counts()

    # Create a pie chart for visual representation
    plt.figure(figsize=(8, 8))
    flavour_counts.plot(kind='pie', autopct='%1.1f%%', startangle=90, colors=plt.cm.Paired.colors)
    plt.title('Foods using Synthetic Aroma Agents Causing Specialized Effects')
    plt.ylabel('')
    plt.tight_layout()
    return plt.gcf()  # Return the figure for embedding

# New function to analyze allergic reactions and plot data
def plot_allergic_reaction_data():
    plt.clf()
    allergic_data = df[df['Causes/Effects'].str.contains('Allergic Reaction', case=False)]

    synthetic_count = len(allergic_data)
    print(f'Number of Synthetic Aroma Agents Causing Allergic Reaction: {synthetic_count}')

    flavours = allergic_data['Flavour of Aroma'].unique()
    print("\nFlavours of Synthetic Aroma Agents Causing Allergic Reaction:")
    print(flavours)

    print("\nNatural replacements of Synthetic Aroma Agents Causing Allergic Reaction:")
    replacement_table = allergic_data[['Synthetic Aroma Agent Name', 'Natural Agents']]
    print(replacement_table)

    flavour_counts = allergic_data['Food'].value_counts()

    plt.figure(figsize=(8, 8))
    flavour_counts.plot(kind='pie', autopct='%1.1f%%', startangle=90, colors=plt.cm.Paired.colors)
    plt.title('Foods using Synthetic Aroma Agents Causing Allergic Reaction')
    plt.ylabel('')
    plt.tight_layout()
    return plt.gcf()  # Return the figure for embedding

# Map function names to analytical questions
graph_funcs = {
    "What is the Distribution of Effects of Synthetic Aromatic Agents?": plot_effect_distribution,
    "What are the Effects by Flavour of Aroma?": plot_flavor_effect_heatmap,
    "What are the Effects by Food Type?": plot_food_effect_heatmap,
    "What are the Foods using Synthetic Aroma Agents Causing Irritation?": plot_irritation_effects,
    "What are the Foods using Synthetic Aroma Agents Causing Specialized Effects?": plot_specialized_effects,
    "What are the Foods using Synthetic Aroma Agents Causing Allergic Reaction?": plot_allergic_reaction_data,
    "What are the Foods using Synthetic Aroma Agents Causing Systemic Toxicity?": lambda: analyze_systemic_toxicity(df),
    "What are the Foods using Synthetic Aroma Agents with Safety Approvals?": lambda: analyze_safety_approvals(df)  # New entry
}

# Create Tkinter GUI
root = tk.Tk()
root.title("Perfume Data Analytics Dashboard")
root.geometry("1000x700")

label = ttk.Label(root, text="Select a Question:")
label.pack(pady=10)

combo = ttk.Combobox(root, values=list(graph_funcs.keys()), state="readonly", width=60)
combo.pack()

frame = tk.Frame(root)
frame.pack(fill=tk.BOTH, expand=True)

canvas_widgets = []  # List to hold canvas widgets

def show_graph(event=None):
    global canvas_widgets
    # Destroy existing canvas widgets
    for canvas in canvas_widgets:
        canvas.get_tk_widget().destroy()
    canvas_widgets.clear()  # Clear the list

    question = combo.get()
    if question in ["What are the Effects by Food Type?"]:
        figs = graph_funcs[question]()  # Get both figures
        for fig in figs:
            # Create a canvas for each figure
            canvas_widget = FigureCanvasTkAgg(fig, master=frame)
            canvas_widget.draw()
            canvas_widget.get_tk_widget().pack(fill=tk.BOTH, expand=True)
            canvas_widgets.append(canvas_widget)  # Store the canvas widget
    else:
        fig = graph_funcs[question]()  # Get single figure
        canvas_widget = FigureCanvasTkAgg(fig, master=frame)
        canvas_widget.draw()
        canvas_widget.get_tk_widget().pack(fill=tk.BOTH, expand=True)
        canvas_widgets.append(canvas_widget)  # Store the canvas widget

combo.bind("<<ComboboxSelected>>", show_graph)

# Show the first graph by default
combo.current(0)
show_graph()

root.mainloop()
