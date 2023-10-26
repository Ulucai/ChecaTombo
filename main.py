import PySimpleGUI as sg
from scan_planilha import Planilha
import os.path
import pyperclip

def create_window(antiga:Planilha, nova:Planilha):            
    data = antiga.compara(nova)
    table_viewer_column = [

        [sg.Text("Listas para Comparação")],

        [sg.Table(values=data, headings=["Bens que saíram", "Bens que entraram"],
                auto_size_columns= True, display_row_numbers=False, expand_y=True,
                select_mode=sg.TABLE_SELECT_MODE_EXTENDED, enable_events=True, key="-TABLE-")],
    ]        
    
    layout = [
        [
            sg.Column(table_viewer_column, expand_x=True, expand_y=True),                   
        ]
    ]
    window = sg.Window("Relatório", layout, finalize=True, resizable=True)
    window.bind("<Control-C>", "Control-C")
    window.bind("<Control-c>", "Control-C")
    return window, data

def main():
    data = []
    # Janela Inicial     
    left_file_column = [

        
        [sg.Text("Planilha Antiga")],

        [sg.In(size=(25, 1), enable_events=True, key="-FILE1-"),
        # Limita a visualização para arquivos excel
        sg.FileBrowse(file_types=[("Arquivos Excel", "*.xls *.xlsx")])],
                            
    ]

    right_file_column =[
        
        [sg.Text("Planilha Nova")],

        [sg.In(size=(25, 1), enable_events=True, key="-FILE2-"),            
        # Limita a visualização para arquivos excel
        sg.FileBrowse(file_types=[("Arquivos Excel", "*.xls *.xlsx")])],
        
    ]

    bottom = [
        [sg.Button(button_text="Comparar",key="-OK-")],
    ]



    
# ----- Layout  -----
    layout = [

        [            
            [sg.Frame('Antiga', layout=left_file_column, size=(400, 100), vertical_alignment='top')],
            
            [sg.Frame('Nova', layout=right_file_column, size=(400, 100), vertical_alignment='top')],
            
            [sg.HSeparator()],

            sg.Button(button_text="Comparar",key="-OK-"),            
        ]

    ]


    main_window = sg.Window("Comparador", layout, finalize=True)
    window_report = None

    # Loop para leitura de eventos

    while True:

        window, event, values = sg.read_all_windows()
        
        if event == "Exit" or event == sg.WIN_CLOSED:
            window.close()
            if window == window_report:
                window_report = None
            elif window == main_window:
                break

        
        elif event == "-OK-":
            # Verifica se dois arquivos foram selecionados e armazena o path
            try:
                file1 = values["-FILE1-"]
                file2 = values["-FILE2-"]                
                
                file_list = [file1, file2]

            except:

                file_list = []


            fnames = [

                f

                for f in file_list

                if os.path.isfile(f)
                
                and f.lower().endswith((".xls", ".xlsx"))
            ]            

            if not window_report and len(fnames)==2:
                cols=[1,3,14]
                antiga = Planilha(fnames[0], require_cols=cols)
                nova = Planilha(fnames[1], require_cols=cols)
                window_report, data = create_window(antiga, nova)               
        if event == "Control-C":                        
            items = values['-TABLE-']            
            lst = list(map(lambda x:' '.join(data[x]), items))
            text = "\n".join(lst)
            pyperclip.copy(text)            
    window.close()
if __name__ == "__main__":
    main()