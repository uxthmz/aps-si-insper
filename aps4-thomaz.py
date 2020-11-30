# APS4 (Rodrigo Souza, Thomaz Edmundo, Pedro Queiroz, Vitor Orsi)
import xlrd
import numpy as np
import matplotlib.pyplot as fig
import statistics as st
import math as mt

wb=xlrd.open_workbook("APS4.xlsx")
plan=wb.sheet_by_name("Planilha1")

x=plan.col_values(0)
y=plan.col_values(1)

n=plan.nrows
m=plan.ncols

v=np.array(y)

retorno=(v[1:n]-v[0:n-1])/v[0:n-1]

fig.subplot(221)
fig.plot(y)
fig.title("PREÇO")
fig.xlabel("Dias")
fig.ylabel("Fechamento")
fig.grid()

fig.subplot(222)
fig.plot(retorno)
fig.title("RETORNO")
fig.xlabel("Dias")
fig.ylabel("Retorno")
fig.grid()

fig.subplot(223)
fig.hist(retorno,bins=5,density=True)
fig.title("HISTOGRAMA")
fig.xlabel("Preço")
fig.ylabel("Frequencia (%)")
fig.grid()

fig.subplot(224)
fig.boxplot(retorno)
fig.title("BOXPLOT")
fig.grid()

mediana=st.median(retorno)
desvio=st.pstdev(retorno)
media=st.mean(retorno)
maximo=retorno.max()
minimo=min(retorno)
otimista=media+2*desvio/mt.sqrt(n)
pessimista=media-2*desvio/mt.sqrt(n)

print("----------------------------------")
print("Mediana =" ,mediana)
print("Desvio =" ,desvio)
print("Media = " ,media)
print("Maximo = " ,maximo)
print("Minimo = " ,minimo)
print("Intervalo de confianca = " ,otimista, "," ,pessimista)
print("----------------------------------")
