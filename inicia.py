
import numpy as np
import pandas as pd
import re
import os
import locale
locale.setlocale(locale.LC_ALL, 'pt_BR.UTF-8')

#https://docs.google.com/spreadsheets/d/1N6JFMIQR71HF5u5zkWthqbgpA8WYz_0ufDGadeJnhlo/edit#gid=0
#1N6JFMIQR71HF5u5zkWthqbgpA8WYz_0ufDGadeJnhlo

def listTelefones(telefones):
    telefones_conv = []

    for tel in telefones:
        # print(tel)
        telp = re.sub('[^0-9]', '', str(tel))  # refiltro.findall(str(tel))

        t = len(telp)
        if t == 0:
            resp = ""
        elif t == 12:
            resp = '+' + telp[:2] + ' ' + telp[2:12]
        else:
            resp = '+55 ' + telp
        telefone = str(resp)
        telefones_conv.append(telefone)

    return  telefones_conv

def formatarCamposValorDesconto(valores, d):
    i = 0

    valor_com_desconto = []
    valores_conv = []

    for val in valores:

        val = re.sub('[^0-9^.^,]', '', str(val))

        _val = str(val)

        if _val.find('.') > -1:
            resp = _val
        elif _val.find(',') > -1:
            resp = _val.replace(",", ".")
        else:
            valp = _val

            t = len(valp)

            if t == 0:
                resp = 0
            elif t == 3:
                resp = valp[:2] + "." + valp[2:] + "00"
            elif t == 4:
                resp = valp[:2] + "." + valp[2:] + "00"
            else:
                resp = valp[:3] + "." + valp[3:] + "00"

        desconto = str(d[i])

        if not desconto.isdigit():
            desconto = 0

        i += 1

        valor = float(resp)
        desc = valor - (valor * float(desconto) / 100.0)

        valor = locale.currency(valor, grouping=True, symbol=True)
        valores_conv.append(valor)

        desc = locale.currency(desc, grouping=True, symbol=True)
        valor_com_desconto.append(desc)

    return valores_conv, valor_com_desconto


if __name__ == '__main__':

    arquivo = 'Teste importador.xlsx'

    print("Selecione: ")
    print("1) google sheets")
    print("2) do arquivo")

    sel = input()

    if sel == '1':

        import googlesheet

        usu = googlesheet.getPlanilhaGoogle('1N6JFMIQR71HF5u5zkWthqbgpA8WYz_0ufDGadeJnhlo', 'usuarios')
        dep = googlesheet.getPlanilhaGoogle('1N6JFMIQR71HF5u5zkWthqbgpA8WYz_0ufDGadeJnhlo', 'dependentes')

        print("Preparando dados...")

        header_usu = usu[0]
        header_dep = dep[0]

        del(usu[0])
        del(dep[0])

        usuarios = pd.DataFrame(usu, columns = header_usu)
        dependentes = pd.DataFrame(dep, columns = header_dep)

        #telefones = []
        #valores = []
        #desporc = []

        #for linha in usu:
        #   telefones.append(linha[3])
        #   valores.append(linha[4])
        #   desporc.append(linha[5])

        print("Formatando campo telefone...")
        # formatar campo telefone - 1) deixar so numeros; 2) normalizar com mascara
        telefones = usuarios['telefone']
        telefones_conv = listTelefones(telefones)

        print("Preparando campo valor...")
        # formatar campo valor
        valores = usuarios['valor']
        desporc = usuarios['desconto']
        valores_conv, valor_com_desconto = formatarCamposValorDesconto(valores, desporc)


    else:

        if not os.path.exists(arquivo):
            print("Arquivo '%s' não encontroado." % arquivo)
            quit()

        print("Inicializando tabelas...")
        usuarios = pd.read_excel(arquivo, sheet_name='usuarios', dtype={'telefone': str, 'valor': str})
        dependentes = pd.read_excel(arquivo, 'dependentes')

        print("Formatando campo telefone...")
        # formatar campo telefone - 1) deixar so numeros; 2) normalizar com mascara

        telefones = usuarios['telefone']
        telefones_conv = listTelefones(telefones)

        print("Preparando campo valor...")
        # formatar campo valor

        valores = usuarios['valor']
        desporc = usuarios['desconto']

        valores_conv, valor_com_desconto = formatarCamposValorDesconto(valores, desporc)


    print("Criando DataFrame Usuarios...")
    usuarios_form = pd.DataFrame({'id': usuarios['id'], 'nome': usuarios['nome'],
                                  'email': usuarios['email'], 'telefone': telefones_conv,
                                  'valor_total': valores_conv, 'valor_com_desconto': valor_com_desconto})

    usuarios_form['telefone'] = usuarios_form['telefone'].astype('str')

    print("Criando DataFrame Dependentes...")
    dependentes_form = pd.DataFrame({'id': dependentes['id'], 'usuario_id': dependentes['user_id'],
                                     'dependente_de_id': dependentes['dependente_id'],
                                     'data': dependentes['data_hora']})

    print("Salvando Usuarios .csv ")
    usuarios_form.to_csv('usuarios.csv', sep=";", columns=['id', 'nome',
                                                           'email', 'telefone', 'valor_total', 'valor_com_desconto'],
                         index=False, index_label=False)
    print("Salvando Dependentes .csv ")
    dependentes_form.to_csv('dependentes.csv', sep=";", index=False, index_label=False)


    print("Finalizado.")






