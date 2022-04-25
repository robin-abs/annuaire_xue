import sys
import os
import pandas as pd


def new_contact_row(df_in,
                    i,
                    base_columns,
                    competences_colums,
                    self=True,
                    col_suffix=''):
    new_row = {}
    for c in base_columns:
        new_row[c] = df_in.iloc[i][c + col_suffix]
    competences = []
    for c in competences_colums:
        if df_in.iloc[i][c + col_suffix] == 'X':
            competences.append(c)
    new_row['Domaines de compétence'] = ', '.join(competences)
    if self:
        new_row['Contact Principal (Nom)'] = 'Lui-même'
        new_row['Contact Principal (Prénom)'] = 'Lui-même'
    else:
        new_row['Contact Principal (Nom)'] = df_in.iloc[i]['Nom']
        new_row['Contact Principal (Prénom)'] = df_in.iloc[i]['Prénom']
    new_row['Contact Principal (mail)'] = df_in.iloc[i]['Adresse mail']
    new_row['Contact Principal (téléphone)'] = df_in.iloc[i]['Numéro de téléphone']

    return new_row


def adjust_xls_cells_format(writer, df, sheet_name='TOUS'):
    header_props = {
        'bg_color': '#449dad',
        'font_color': '#ffffff',
        'text_wrap': True,
        'align': 'center',
        'bold': True,
        'border': 1
    }
    self_props = {
        'bold': True,
    }

    # header & self cells format
    header_format = writer.book.add_format(header_props)
    self_format = writer.book.add_format(self_props)

    # Set cells format
    for col_num, value in enumerate(df.columns.values):
        writer.sheets[sheet_name].set_column(col_num, col_num, 23)
        writer.sheets[sheet_name].write(0, col_num, value, header_format)
        # Put main contact in Bold
        for i in df.index:
            if df.loc[i, "Contact Principal (Nom)"] == "Lui-même":
                if pd.notnull(df.loc[i, value]):
                    writer.sheets[sheet_name].write(i+1, col_num, df.loc[i, value], self_format)

    # Set header height
    writer.sheets[sheet_name].set_row(0, 25)

    # Set auto-filter
    (max_row, max_col) = df.shape
    writer.sheets[sheet_name].autofilter(0, 0, max_row, max_col - 1)

    return writer


def export(df_in, N, base_columns, competences_colums):
    # Sheet with all contacts
    df_out_all = pd.DataFrame(columns=['Nom', 'Prénom', 'Structure & fonction', 'Courte description de ses engagements',
                                       'Site personnel (LinkedIn ou autre)', 'Domaines de compétence',
                                       'Contact Principal (Nom)', 'Contact Principal (Prénom)',
                                       'Contact Principal (mail)', 'Contact Principal (téléphone)'],
                              )

    for i in range(len(df_in)):
        # Contact Principal
        row_main_temp = new_contact_row(df_in, i, base_columns=base_columns, competences_colums=competences_colums)
        df_out_all = df_out_all.append(row_main_temp, ignore_index=True)

        # Contacts Secondaires
        for j in range(N):
            if pd.isnull(df_in.iloc[i]['Nom' + f'.{j + 1}']):
                break
            row_sec_temp = new_contact_row(df_in, i, base_columns=base_columns, competences_colums=competences_colums,
                                           self=False, col_suffix=f'.{j + 1}')
            df_out_all = df_out_all.append(row_sec_temp, ignore_index=True)

    df_out_all = df_out_all.drop_duplicates().reset_index(drop=True)

    # Sheet per competence field
    d_competences_sheets = {}
    for competence in competences_colums:
        d_competences_sheets[competence] = df_out_all[
            df_out_all['Domaines de compétence'].apply(lambda x: competence in x)].reset_index(drop=True).drop(
            columns='Domaines de compétence')

    # Export to xls file
    writer = pd.ExcelWriter('annuaire_x_urgence_ecologique.xlsx')
    df_out_all.to_excel(writer, index=False, sheet_name='TOUS', freeze_panes=(1,0))
    writer = adjust_xls_cells_format(writer, df_out_all)

    for competence in competences_colums:
        sheet_name = competence.replace('/', '-').replace("Impact environnemental de l'industrie",
                                                          "Impact env. de l'industrie")
        d_competences_sheets[competence].to_excel(writer,
                                                  index=False,
                                                  sheet_name=sheet_name,
                                                  freeze_panes=(1,0))
        writer = adjust_xls_cells_format(writer, d_competences_sheets[competence], sheet_name=sheet_name)

    writer.save()


if __name__ == "__main__":
    file_name = 'annuaire_x_urgence_ecologique.csv'
    if len(sys.argv)>1:
        file_name = sys.argv[1]
    print(f'Import du csv {file_name}')
    path = os.getcwd()+'/'+file_name
    df_in = pd.read_csv(path, skiprows=2, sep=';', low_memory=False)
    N = int(df_in.columns[-1][-1])

    base_columns = ['Nom', 'Prénom', 'Structure & fonction',
                    'Courte description de ses engagements', 'Site personnel (LinkedIn ou autre)']
    competences_colums = ['Agriculture',
                          'Aménagement/construction', 'Biodiversité', 'Déchets', 'Eau', 'Énergie',
                          'Économie', 'Gouvernance', "Impact environnemental de l'industrie",
                          'Justice environnementale', 'Justice sociale', 'Low-Tech',
                          'Philosophie', "Qualité de l'air", 'Transport/mobilité']

    print(f'Préparation du fichier excel...')
    export(df_in, N, base_columns, competences_colums)
    print('Fichier excel prêt !')
