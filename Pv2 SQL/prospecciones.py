
"""
p1:
listado de empresas que se enrollean


P1 (parte generica)


* creo un df con:
nombre, dominio, región, industria, fecha de creación y de actualización (en este caso
serian iguales)

* importo la table de empresas de Prospecciones Unificadas como df
* importo la table de empresas de Hubspot como df

# por empresa me refiero al id => empresa una combinacion del dominio y el pais

for df in dfs:
    crear columna que sea dominio + pais
    #sirve como dominio de la empresa


---
P1_empresas_nuevas:
---
* creamos una nueva df con los siguientes
for empresa in df
    if empresa in table de empresas de PU or empresa in table de empresas de Hubspot
        continue
    agregar empresa a nuevo_df
se carga a PU nuevo_df (agregandolo al fondo)

Y la usamos para generar la tabla fdm

---
P1_empresas_listado:
---
* creamos una nueva df con los siguientes
for empresa in df
    if empresa in table de empresas de PU or empresa in table de empresas de Hubspot
        update empresa creat date  al que esta en huspot
        agregar empresa a nuevo_df_1
        continue
    agregar empresa a nuevo_df_2

se agrega a PU nuevo_df_2
y se agrega/ actualiza a PU nuevo_df_1

---
P1_prioridad:
---


"""

