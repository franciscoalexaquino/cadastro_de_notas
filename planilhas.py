import win32com.client as win32
import streamlit as st
import pythoncom
import tempfile
import pandas as pd

# Removendo espaços do topo da página
st.markdown("""
<style>
section.stMain .block-container {
    padding-top: 0rem;
    z-index: 1;
}
</style>""", unsafe_allow_html=True)


with st.container(border=True, width="stretch"):
    container_geral = st.container()
    container_1 = st.container(border=True)
    powerpoint_file, notas_file = st.columns(2)
    container_2 = st.container(border=True)

container_geral.title("CADASTRO DE NOTAS")
bimestre_desejado = container_1.text_input('EM QUAL BIMESTRE DESEJA ISERIR AS NOTAS?').strip()
container_1.info("Para inserir no 1° bimestre, digite o número 1. Para o 2° bimestre, " \
"digite 2. Para o 3° bimestre, digite 3. Para o 4° bimestre, digite 4.")

btn_enviar = container_2.button('INSERIR NOTAS')

arquivo_pptx = powerpoint_file.file_uploader("SELECIONE O SEU SLIDE")
arquivo_notas = notas_file.file_uploader("SELECIONE A PLANILHA DO PROFESSOR ONLINE")

def buscar_notas():
    pass

def inserir_notas(path, bimestre):
    if bimestre:
        # Inicializa o COM / Ela é necessária em alguns contextos
        pythoncom.CoInitialize()

        # Abre o PowerPoint
        ppt = win32.gencache.EnsureDispatch("PowerPoint.Application")
        ppt.Visible = True

        # Caminho da apresentação
        presentation = ppt.Presentations.Open(path, ReadOnly=False, WithWindow=True)

        # Garante que está em modo "slide view"
        ppt.ActiveWindow.ViewType = 9

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
                    ultima_linha = ws.Cells(ws.Rows.Count, int(bimestre) + 1).End(-4162).Row

                    # Preenche a coluna 4 (D) com valor 2
                    for i in range(3, ultima_linha):
                        ws.Cells(i, int(bimestre) + 1).Value = 3

                    print(f"Planilha do slide {slide.SlideIndex} editada com sucesso!")
                except Exception as e:
                    print(f"Erro ao editar planilha no slide {slide.SlideIndex}: {e}")
            else:
                print(f"Nenhuma planilha encontrada no slide {slide.SlideIndex}.")

        # Salva e fecha
        presentation.Save()
        presentation.Close()
        ppt.Quit()
        return container_geral.success("NOTAS INSERIDAS")
    else:
        return container_geral.warning("SELECIONE UM BIMESTRE!")


if btn_enviar:
    if arquivo_pptx is not None:
        with tempfile.NamedTemporaryFile(delete=False, suffix=".pptx") as tmp:
            tmp.write(arquivo_pptx.getbuffer())
            path = tmp.name
            print(path)

        inserir_notas(path, bimestre_desejado)
    else:
        container_geral.warning("SELECIONE A PLANILHA PARA CALCULAR AS NOTAS")

