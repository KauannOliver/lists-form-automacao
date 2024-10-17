import tkinter as tk
import threading
import functions as f  ### importa as funções do arquivo functions.py ###

### função para iniciar o preenchimento do formulário ###
def iniciar_preenchimento():
    if f.driver is None:
        print("driver não foi iniciado. por favor, clique em 'abrir navegador' primeiro.")  ### mensagem de alerta se o driver não foi iniciado ###
    else:
        threading.Thread(target=f.preencher_formulario).start()  ### inicia o preenchimento em uma nova thread ###

### criando a interface gráfica com tkinter ###
def iniciar_tela_tkinter():
    root = tk.Tk()  ### cria a janela principal do tkinter ###
    root.title("automação de formulário")  ### define o título da janela ###
    root.geometry("300x200")  ### define o tamanho da janela ###

    ### cria um botão para abrir o driver e acessar o link ###
    abrir_driver_button = tk.Button(root, text="abrir navegador", command=lambda: threading.Thread(target=f.abrir_driver).start(), font=("Arial", 12), bg="blue", fg="white")
    abrir_driver_button.pack(pady=20)  ### posiciona o botão na janela ###

    ### cria um botão que inicia o processo de preenchimento ###
    start_button = tk.Button(root, text="inserir dados", command=iniciar_preenchimento, font=("Arial", 12), bg="green", fg="white")
    start_button.pack(pady=20)  ### posiciona o botão na janela ###

    root.mainloop()  ### inicia o loop principal da interface gráfica ###

### função principal que inicia a interface tkinter ###
if __name__ == "__main__":
    iniciar_tela_tkinter()
