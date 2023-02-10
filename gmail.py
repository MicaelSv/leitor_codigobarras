from imbox import Imbox
from datetime import datetime
import pandas as pd

username = 'micael.msilva012@gmail.com'
password = open('passwords/pass', 'r').read() #aq vai ler a senha q eu peguei da conta do google
host = "imap.gmail.com"

mail = Imbox(host, username = username, password = password, ssl = True) #ssl é pra estabelecer uma conexao segura

mensagens = mail.messages(raw='has:attachment') #attachment vai explorar apenas as msgs que contem anexos
