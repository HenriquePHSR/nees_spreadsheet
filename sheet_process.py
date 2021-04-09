import os
import pandas as pd
import re
import openpyxl
import numpy as np
from datetime import datetime
from dateutil.relativedelta import *

def validate(cpf: str) -> bool:
    numbers = [int(digit) for digit in cpf if digit.isdigit()]

    if len(numbers) != 11 or len(set(numbers)) == 1:
        return False

    sum_of_products = sum(a*b for a, b in zip(numbers[0:9], range(10, 1, -1)))
    expected_digit = (sum_of_products * 10 % 11) % 10
    if numbers[9] != expected_digit:
        return False

    sum_of_products = sum(a*b for a, b in zip(numbers[0:10], range(11, 1, -1)))
    expected_digit = (sum_of_products * 10 % 11) % 10
    if numbers[10] != expected_digit:
        return False

    return True

def autocomplete(str, lenght, char, mode):
    while(len(str) < lenght):
        if mode == 'left':
            str = char + str
        if mode == 'right':
            str = str + char
    return str

def build_df(args):
    if (len(args)!=0):
        df_in = pd.read_excel(args[0])
    else:
        f_name = input()
        df_in = pd.read_excel(f_name)
    
    cursos = sorted(df_in.Curso.unique())
    if (len(cursos) < 11):
        print("# ERRO :: Numero de Cursos inferior ao esperado")
        return 0
    tags = ['ACLF','AF', 'AECP', 'ASS', 'CB', 'CP', 'EDCR', 'EMP', 'FEF', 'IEA', 'PCP']
    cursos_dict = dict(zip(cursos, tags))



    df_out = pd.DataFrame()
    df_out['E-mail'] = pd.Series(df_in['E-mail'], dtype="string")
    df_out['Turma'] = [cursos_dict[x] for x in df_in['Curso'].values]
    df_out['Nome do curso'] = df_in['Curso']
    
    ''' AUTOCOMPLETE CPF  AND FORMAT'''
    df_out['CPF'] = [autocomplete(str(x), 11, '0', 'left') for x in df_in['CPF'].values]
    df_out['CPF'] = pd.Series(df_out['CPF'], dtype="string")

    df_out['Nome'] = pd.Series(df_in['Aluno'], dtype="string")
    df_out['Nome'] = [x.title() for x in df_out['Nome'].values]
    
    df_out['periodo'] = [str((datetime.strptime(x, '%d/%m/%Y')).date())+' a ' + str((datetime.strptime(x, '%d/%m/%Y')+relativedelta(months=+1)).date()) for x in df_in['Data da Prova']]

    df_in['Data da Prova'] = [datetime.strptime(x, '%d/%m/%Y') for x in df_in['Data da Prova']]
    df_in['semestre'] = np.where(df_in['Data da Prova'] < datetime.strptime('01/06/2020', '%d/%m/%Y'), '01', '02')
    df_in['Data da Prova'] = pd.Series(df_in['Data da Prova'], dtype="string")
    df_in['Ano'] = [x.split('-')[0] for x in df_in['Data da Prova'].values]
    df_out['Turma'] = df_out['Turma']+df_in['Ano']+'.'+df_in['semestre']

    df_out['E-mail'] = pd.Series(df_out['E-mail'], dtype="string")
    df_out['Turma'] = pd.Series(df_out['Turma'], dtype="string")
    df_out['Nome do curso'] = pd.Series(df_out['Nome do curso'], dtype="string")
    df_out['CPF'] = pd.Series(df_out['CPF'], dtype="string")
    df_out['Nome'] = pd.Series(df_out['Nome'], dtype="string")
    df_out['periodo'] = pd.Series(df_out['periodo'], dtype="string")
    ''' VALIDATE CPF '''
    df_out_final = df_out[df_out['CPF'].apply(lambda x: validate(str(x)))]

    ''' NOT VALIDE CPFs '''
    df_out_cpf_rejected = df_out[df_out['CPF'].apply(lambda x: not validate(str(x)))]
    if not os.path.exists('output'):
        os.mkdir('output')
    df_out_final.to_excel('output/ouput_df.xlsx')
    print('output file: ouput_df.xlsx')
    df_out_cpf_rejected.to_excel('output/output_df_cpf_rejected.xlsx')
    print('output file: output_df_cpf_rejected.xlsx')

    return df_out_final, df_out_cpf_rejected

def fill_mail(sender_email, receiver_email, attachment, subject, body):
    # Create a multipart message and set headers
    message = MIMEMultipart()
    message["From"] = sender_email
    message["To"] = receiver_email
    message["Subject"] = subject
    message["Bcc"] = receiver_email  # Recommended for mass emails

    # Add body to email
    message.attach(MIMEText(body, "plain"))

    file_path = attachment

    if file_path != '':
        filename = file_path.split("/")[-1]
        # open has a binary stream
        with open(file_path, "rb") as attachment:
            # Add file as application/octet-stream
            # Email client can usually download this automatically as attachment
            part = MIMEBase("application", "octet-stream")
            part.set_payload(attachment.read())

        # Encode file in ASCII characters to send by email
        encoders.encode_base64(part)

        # Add header as key/value pair to attachment part
        part.add_header(
            "Content-Disposition",
            f"attachment; filename= {filename}",
        )
        # Add attachment to message and convert message to string
        message.attach(part)

    text = message.as_string()
    return text


def dispatch_email_from_df(args):
    fname = args[0]
    print(str(fname.split('/')[-1]) + " selecionado")
    try:
        df = pd.read_excel(fname)
        # henrique_pedro@id.uff.br
        # Log in to server using secure context and send email
        username = str(input("email: "))
        password = "njmrtbutmennrekn"
        subject = "An email with attachment from Python"
        body = "This is an email with attachment sent from Python"
        # subject = input("assunto: ")
        # body = input("corpo: ")
        # print('select attachment: ')
        # attachment = fd.askopenfilename()
        if len(args) > 1:
            attachment = args[1]
        else:
            attachment = ''
        user = {
            "username": username,
            "pass": password
        }
        context = ssl.create_default_context()
        with smtplib.SMTP_SSL("smtp.gmail.com", 465, context=context) as server:
            if not is_valid(user["username"]):
                print("Invalid Email")
                raise Exception()
            server.login(user["username"], user['pass'])

            to_send = []
            if 'E-Mail' in df.columns:
                for email in df['E-Mail'].values:
                    if is_valid(email):
                        to_send.append(email)
            else:
                print("email column not found")
                to_send = None
            not_send = []
            for email in to_send:
                try:
                    server.sendmail(user['username'], email,
                                    fill_mail(user['username'], email, attachment, subject, body))
                    print("mail sent to " + email)
                except:
                    print("exception occurred mail to " + email + " not send")
                    not_send.append(email)

    except FileNotFoundError:
        print('File not found')
    except:
        print('Something went wrong')


def is_valid(email):
    if type(email) == str:
        if re.search(r"(^[a-zA-Z0-9_.+-]+@[a-zA-Z0-9-]+\.[a-zA-Z0-9-.]+$)", email):
            return True
        else:
            return False
    else:
        return False