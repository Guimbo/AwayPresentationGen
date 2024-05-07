import tkinter as tk
from tkinter import PhotoImage
from tkinter import messagebox
from presentation_generator import PresentationGenerator

class PresentationGeneratorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Gerador de Apresentações")
        self.root.geometry("800x600")

        # Adicionando imagem de fundo à interface
        background_image = PhotoImage(file="background_image.png")
        background_label = tk.Label(root, image=background_image)
        background_label.place(x=0, y=0, relwidth=1, relheight=1)
        background_label.image = background_image

        # Adicionando logo à interface
        logo_image = PhotoImage(file="logo_image.png")
        logo_label = tk.Label(root, image=logo_image)
        logo_label.pack()

        # Botão "Gerar"
        generate_button = tk.Button(root, text="Gerar", command=self.generate_and_save)
        generate_button.pack()

    def generate_and_save(self):
        # Cria uma instância da classe PresentationGenerator
        presentation_generator = PresentationGenerator(self.get_data(), "nome_arquivo")
        
        # Gera a apresentação
        presentation_generator.generate_presentation()

        # Salva o caminho de destino das apresentações em path.txt
        if presentation_generator.dest_path:
            with open("path.txt", "w") as file:
                file.write(presentation_generator.dest_path)

    def get_data(self):
        # Retorna os dados inseridos pelo usuário
        pass
