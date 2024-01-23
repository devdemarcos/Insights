# Insights
## Análise de dados associado ao operacional de engenharia
import pandas as pd
import matplotlib.pyplot as plt
import win32com.client as win32

# importar a base de dados
ensaio_laboratorio = pd.read_excel('arquivo.xlsx')

# visualizar a base de dados
pd.set_option('display.max_columns', None)
none = ('-')

# caracteristica do material
caracteristica = ensaio_laboratorio[['Resistencia', 'Traço - carrinho']]

# criar gráfico de dispersão
caracteristica.plot(kind='scatter', x='Resistencia', y='Traço - carrinho')
plt.show()


# enviar um email com o relatório
outlook = win32.Dispatch('outlook.application')
mail = outlook.CreateItem(0)
mail.To = 'email@gmail.com'
mail.Subject = 'Plano de ação do x'
mail.HTMLBody = f'''
<p>Prezada Analista,</p>

<p>Segue o Plano de ação dos ensaios laboratoriais feitos com os agregados produzidos na unidade de beneficiamento de resíduos da construção civil.</p>

<p>Caracteristicas da aplicação do material:</p>
{caracteristica.to_html()}
{caracteristica.plot_html()}

<p>Qualquer dúvida estou a disposição.</p>
<p>Marcos Vinicius da Silva, Equipe técnica Grupo x Att.,</p>
<p></p>
'''
mail.Send()
print('Email Enviado')
