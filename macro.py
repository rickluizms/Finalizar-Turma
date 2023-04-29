import time
import win32com.client as xl



def main_convert():
    wb = xl.Dispatch("Excel.Application")
    #Torna visivel/invisivel (2°plano)
    wb.Visible = True

    #Abre o arquivo
    main_convert = wb.Workbooks.open("C:\workspace\FinalizarTurma\Input\Main_conversor.xlsm")

    #Atualiza a tabela | main-convert process
    main_convert.RefreshAll()
    time.sleep(10)

    #Executar Macro
    main_convert.Application.Run("getTermoApp")

    main_convert.Application.Run("getCertificateApp")

    main_convert.Application.Run("getCartaApp")

    main_convert.Save()
    wb.Quit()


def TermoApp():
    wb = xl.Dispatch("Excel.Application")
    #1° - Termo App
    termoApp = wb.Workbooks.open("C:\workspace\FinalizarTurma\App\TermoApp\TermoApp.xlsm")

    termoApp.RefreshAll()
    time.sleep(10)

    termoApp.Application.Run("getTermo")
    termoApp.Save()
    wb.Quit()

def CertificateApp():
    wb = xl.Dispatch("Excel.Application")
    #2° - Certificate App
    certificateApp = wb.Workbooks.open("C:\workspace\FinalizarTurma\App\CertificateApp\CertificateAppGP\CertificateApp.xlsm")

    certificateApp.RefreshAll()
    time.sleep(10)

    certificateApp.Application.Run("getCertificate")
    certificateApp.Save()
    wb.Quit()

def CartaApp():
    wb = xl.Dispatch("Excel.Application")
    #3° - Carta App
    cartaApp = wb.Workbooks.open("C:\workspace\FinalizarTurma\App\CartaApp\CartaApp.xlsm")

    cartaApp.RefreshAll()
    time.sleep(10)

    cartaApp.Application.Run("getCarta")
    cartaApp.Save()
    wb.Quit()

def ListaApp():
    wb = xl.Dispatch("Excel.Application")
    #4° - Lista App
    listaApp = wb.Workbooks.open("C:\workspace\FinalizarTurma\App\ListaApp\ListaApp.xlsm")

    listaApp.RefreshAll()
    time.sleep(10)

    listaApp.Application.Run("getListaCertificado")
    listaApp.Application.Run("getListaTermo")
    listaApp.Application.Run("getListaAtividade")
    listaApp.Save()
    wb.Quit()