import win32com.client as win32
import streamlit as st
import pythoncom


# Inicializa o COM
pythoncom.CoInitialize()

# Abre o PowerPoint
ppt = win32.gencache.EnsureDispatch("PowerPoint.Application")
ppt.Visible = True

# Caminho da apresentação
presentation = ppt.Presentations.Open(r"C:\Users\AdminUser\Documents\projeto_2\notas.pptx", ReadOnly=False, WithWindow=False)

# Garante que está em modo "slide view"
ppt.ActiveWindow.ViewType = 9

bimestre_desejado = st.text_input('Qual bimestre deseja inserir as notas? (1, 2, 3, 4)').strip()
btn_enviar = st.button('INSERIR NOTAS')


if btn_enviar:
    # Percorre todos os slides
    for slide in presentation.Slides:
        ppt.ActiveWindow.View.GotoSlide(slide.SlideIndex)

        ole_object = None
        for shape in slide.Shapes:
            if shape.Type == 7:  # 7 = OLE Object
                try:
                    shape.OLEFormat.DoVerb()  # ativa o objeto
                    ole_object = shape.OLEFormat.Object
                    break
                except Exception as e:
                    print(f"Erro ao ativar OLE no slide {slide.SlideIndex}: {e}")

        if ole_object:
            try:
                ws = ole_object.Worksheets(1)  # primeira aba da planilha

                # Descobre a última linha da coluna 4 (C)
                ultima_linha = ws.Cells(ws.Rows.Count, int(bimestre_desejado + 1)).End(-4162).Row

                # Preenche a coluna 4 (D) com valor 2
                for i in range(3, ultima_linha):
                    ws.Cells(i, int(bimestre_desejado + 1)).Value = 2

                print(f"Planilha do slide {slide.SlideIndex} editada com sucesso!")
            except Exception as e:
                print(f"Erro ao editar planilha no slide {slide.SlideIndex}: {e}")
        else:
            print(f"Nenhuma planilha encontrada no slide {slide.SlideIndex}.")

    # Salva e fecha
    presentation.Save()
    presentation.Close()
    ppt.Quit()
    st.success("NOTAS INSERIDAS")

else:
    st.warning("INFORME UM BIMESTRE")
