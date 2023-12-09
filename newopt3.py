import tkinter as tk
from tkinter import filedialog
import pandas as pd
import matplotlib.pyplot as plt
import numpy as np

def browse_file():
    file_path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx")])
    if file_path:
        global df
        df = pd.read_excel(file_path, sheet_name=None)
        sheet_names = list(df.keys())
        batch_dropdown.set('')
        batch_menu['menu'].delete(0, 'end') 

        for name in sheet_names:
            batch_menu['menu'].add_command(label=name, command=tk._setit(batch_dropdown, name))

      
        common_subject_menu['menu'].delete(0, 'end')  
        common_subjects = get_common_subjects()
        for subject in common_subjects:
            common_subject_menu['menu'].add_command(label=subject, command=tk._setit(common_subject_dropdown, subject))    



def plot_faculty_member_wise_average_score():
    sheet_name = batch_dropdown.get()
    batch_df = df[sheet_name]
    grouped_df = batch_df.groupby('Faculty Member')['Sem Total'].mean().reset_index()
    plt.figure(figsize=(12, 6))  # Adjust the figure size as per your preference
    
    # Define colors for the bars
    colors = plt.cm.Set3(np.linspace(0, 1, len(grouped_df)))
    
    
    plt.bar(grouped_df['Faculty Member'], grouped_df['Sem Total'], color=colors)
    plt.xlabel('Faculty Member')
    plt.ylabel('Average Score (in Percentage)')
    plt.title('Faculty Wise Average Score')
    plt.xticks(rotation=45, ha='right')  
    
    
    for i, value in enumerate(grouped_df['Sem Total']):
        plt.text(i, value + 1, f"{value:.2f}", ha='center', va='bottom')
    
    plt.tight_layout()  
    plt.show()

def plot_male_vs_female_representation():
    sheet_name = batch_dropdown.get()
    batch_df = df[sheet_name]
    grouped_df = batch_df.groupby('Gender').size().reset_index(name='Count')
    colors = ['pink', 'blue']
    plt.bar(grouped_df['Gender'], grouped_df['Count'], color=colors)
    plt.xlabel('Gender')
    plt.ylabel('Count(in Number)')
    plt.title('Male vs Female Representation')
    for i, count in enumerate(grouped_df['Count']):
        plt.text(i, count + 0.5, str(count), ha='center', va='bottom')
    plt.show()

def plot_average_score_male_vs_female():
    sheet_name = batch_dropdown.get()
    batch_df = df[sheet_name]
    male_avg = batch_df[batch_df['Gender'] == 'M']['Sem Total'].mean()
    female_avg = batch_df[batch_df['Gender'] == 'F']['Sem Total'].mean()
    colors = ['yellow', 'red']

    
    plt.bar(['Male', 'Female'], [male_avg, female_avg], color=colors)
    plt.xlabel('Gender')
    plt.ylabel('Average Score(in Percentage)')
    plt.title('Average Score: Male vs Female')
    for i, score in enumerate([male_avg, female_avg]):
        plt.text(i, score + 1, f"{score:.2f}", ha='center', va='bottom')

    plt.show()
    

def plot_subject_wise_average_score():
    sheet_name = batch_dropdown.get()
    batch_df = df[sheet_name]
    grouped_df = batch_df.groupby(['Semester', 'Subject'])['Sem Total'].mean().reset_index()
    plt.figure(figsize=(10, 6))  
    x_labels = [f"{sem} - {subject}" for sem, subject in zip(grouped_df['Semester'], grouped_df['Subject'])]
    
    
    colors = plt.cm.Set3(np.linspace(0, 1, len(grouped_df)))
    
    
    plt.bar(x_labels, grouped_df['Sem Total'], color=colors)
    plt.xlabel('Semester - Subject')
    plt.ylabel('Average Score (in Percentage)')
    plt.title('Subject Wise Average Score')
    plt.xticks(rotation=45, ha='right')  
    
    for i, value in enumerate(grouped_df['Sem Total']):
        plt.text(i, value + 1, f"{value:.2f}", ha='center', va='bottom')
    
    plt.tight_layout()  
    plt.show()



def plot_section_wise_average_score():
    sheet_name = batch_dropdown.get()
    batch_df = df[sheet_name]
    grouped_df = batch_df.groupby(['Semester', 'Section'])['Sem Total'].mean().reset_index()
    grouped_df['Sem-Section'] = 'Semester ' + grouped_df['Semester'].astype(str) + ' - Section ' + grouped_df['Section'].astype(str)
    plt.figure(figsize=(10, 6))
    
    
    colors = plt.cm.Set3(np.linspace(0, 1, len(grouped_df)))
    
   
    plt.bar(grouped_df['Sem-Section'], grouped_df['Sem Total'], color=colors)
    plt.xlabel('Semester - Section')
    plt.ylabel('Average Score (in Percentage)')
    plt.title('Section Wise Average Score')
    plt.xticks(rotation=45, ha='right')
    
    
    for i, value in enumerate(grouped_df['Sem Total']):
        plt.text(i, value, str(round(value, 2)), ha='center', va='bottom')
    
    plt.tight_layout() 
    plt.show()


def plot_batch_wise_average_score():
    plt.figure(figsize=(10, 6))
    for sheet_name, batch_df in df.items():
        avg_score = batch_df['Sem Total'].mean()
        plt.bar(sheet_name, avg_score)

    plt.xlabel('Batch')
    plt.ylabel('Average Score(in Percentage)')
    plt.title('Batch Wise Average Score')
    plt.xticks(rotation=45, ha='right')
    
    # Add average score values above each bar
    for i, (sheet_name, batch_df) in enumerate(df.items()):
        avg_score = batch_df['Sem Total'].mean()
        plt.text(i, avg_score + 1, f"{avg_score:.2f}", ha='center', va='bottom')

    plt.tight_layout()
    plt.show()

def plot_batch_wise_male_vs_female_representation():
    plt.figure(figsize=(10, 6))
    male_counts = []
    female_counts = []
    
    for sheet_name, batch_df in df.items():
        male_count = batch_df[batch_df['Gender'] == 'M'].shape[0]
        female_count = batch_df[batch_df['Gender'] == 'F'].shape[0]
        male_counts.append(male_count)
        female_counts.append(female_count)

    x_labels = list(df.keys())
    x_ticks = range(len(x_labels))
    width = 0.35

    plt.bar(x_ticks, male_counts, width, label='Female')
    plt.bar(x_ticks, female_counts, width, bottom=male_counts, label='Male')

    plt.xlabel('Batch')
    plt.ylabel('Count (in Number)')
    plt.title('Batch Wise Male vs Female Representation')
    plt.xticks(x_ticks, x_labels, rotation=45, ha='right')
    plt.legend()

   
    for i, (male_count, female_count) in enumerate(zip(male_counts, female_counts)):
        plt.text(i, male_count + female_count, f"Total: {male_count + female_count}", ha='center', va='bottom')
        plt.text(i, male_count + female_count/2, f"Male: {male_count}", ha='center', va='center')
        plt.text(i, male_count/2, f"Female: {female_count}", ha='center', va='center')

    plt.tight_layout()
    plt.show()


def get_common_subjects():
    sheet_names = list(df.keys())
    subjects_set = set(df[sheet_names[0]]['Subject'])

    for name in sheet_names[1:]:
        subjects_set = subjects_set.intersection(set(df[name]['Subject']))

    common_subjects = list(subjects_set)
    return common_subjects

def plot_common_subject_batch_wise():
    subject = common_subject_dropdown.get()
    plt.figure(figsize=(10, 6))
    for sheet_name, batch_df in df.items():
        subject_df = batch_df[batch_df['Subject'] == subject]
        avg_score = subject_df['Sem Total'].mean()
        plt.bar(sheet_name, avg_score)

    plt.xlabel('Batch')
    plt.ylabel('Average Score(in Percentage)')
    plt.title(f'Batch Wise Average Score for Subject: {subject}')
    plt.xticks(rotation=45, ha='right')

    
    
    for i, (sheet_name, batch_df) in enumerate(df.items()):
        subject_df = batch_df[batch_df['Subject'] == subject]
        avg_score = subject_df['Sem Total'].mean()
        plt.text(i, avg_score + 1, f"{avg_score:.2f}", ha='center', va='bottom')

    plt.tight_layout()
    plt.show()




def create_gui():
    global batch_dropdown, batch_menu, df
    root = tk.Tk()
    root.geometry("800x600")
    root.title('Student Data Visualization')
    global common_subject_dropdown, common_subject_menu

    select_file_label = tk.Label(root, text='BROWSE FILE:', font=('New Times Roman', 10, 'bold','underline'))
    select_file_label.pack(pady=10)

    select_file_label = tk.Label(root, text='Select File From Computer:', font=('New Times Roman', 8))
    select_file_label.pack(pady=10)

    browse_button = tk.Button(root, text='BROWSE EXCEL FILE:', command=browse_file)
    browse_button.configure(width=20, height=2, bg='light blue', fg='black')
    browse_button.pack()     


       
    batch_label = tk.Label(root, text='SELECT BATCH:', font=('New Times Roman', 10, 'bold','underline'))
    batch_label.pack(anchor='n', pady=10) 

    batch_dropdown = tk.StringVar(root)
    batch_menu = tk.OptionMenu(root, batch_dropdown, '')
    batch_menu.configure(width = 20,height=1,bg='pink', fg='black')
    batch_menu.pack(anchor='center')

    plot_button1 = tk.Button(root, text='Plot Faculty Member Wise Average Score', command=plot_faculty_member_wise_average_score)
    plot_button1.configure(width=40,height=1, bg='light yellow', fg='black', font=('New Times Roman', 8, 'bold'))
    plot_button1.pack(pady=5)

    plot_button2 = tk.Button(root, text='Plot Male vs Female Representation', command=plot_male_vs_female_representation)
    plot_button2.configure(width=30,height=1, bg='light yellow', fg='black', font=('New Times Roman', 8, 'bold'))
    plot_button2.pack(pady=5)

    plot_button3 = tk.Button(root, text='Plot Average Score: Male vs Female', command=plot_average_score_male_vs_female)
    plot_button3.configure(width=30,height=1, bg='light yellow', fg='black', font=('New Times Roman', 8, 'bold'))
    plot_button3.pack(pady=5)

    plot_button4 = tk.Button(root, text='Plot Subject Wise Average Score', command=plot_subject_wise_average_score)
    plot_button4.configure(width=30,height=1, bg='light yellow', fg='black', font=('New Times Roman', 8, 'bold'))
    plot_button4.pack(pady=5)

    plot_button5 = tk.Button(root, text='Plot Section Wise Average Score', command=plot_section_wise_average_score)
    plot_button5.configure(width=30,height=1, bg='light yellow', fg='black', font=('New Times Roman', 8, 'bold'))
    plot_button5.pack(pady=5)

    batch_analysis_label = tk.Label(root, text='BATCH ANALYSIS:', font=('New Times Roman', 10, 'bold','underline'))
    batch_analysis_label.pack(pady=10)

    plot_button6 = tk.Button(root, text='Plot Batch Wise Average Score', command=plot_batch_wise_average_score)
    plot_button6.configure(width=30,height=1, bg='light yellow', fg='black', font=('New Times Roman', 8, 'bold'))
    plot_button6.pack(pady=5)

    plot_button7 = tk.Button(root, text='Plot Batch Wise Male vs Female Representation', command=plot_batch_wise_male_vs_female_representation)
    plot_button7.configure(width=40,height=1, bg='light yellow', fg='black', font=('New Times Roman', 8, 'bold'))
    plot_button7.pack(pady=5)

    batch_analysis_label = tk.Label(root, text='SUBJECT ANALYSIS:', font=('New Times Roman', 10, 'bold', 'underline'))
    batch_analysis_label.pack(pady=10)

    common_subject_label = tk.Label(root, text='Select Common Subject:')
    common_subject_label.pack()

    common_subject_dropdown = tk.StringVar(root)
    common_subject_menu = tk.OptionMenu(root, common_subject_dropdown, '')
    common_subject_menu.configure(width = 30,height=1,bg='pink', fg='black', font=('New Times Roman', 8, 'bold'))
    common_subject_menu.pack(pady=5)

    plot_button9 = tk.Button(root, text='Plot Batch Wise for Common Subject', command=plot_common_subject_batch_wise)
    plot_button9.configure(width=30,height=1, bg='light yellow', fg='black', font=('New Times Roman', 8, 'bold'))
    plot_button9.pack(pady=5)

    root.mainloop()

if __name__ == '__main__':
    create_gui()
