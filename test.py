def get_teilnehmer_infos(teilnehmer_nummer):
    print('start get teilnehmer infos')
    dataframe1 = pd.read_excel(os.path.abspath(".") + '/files/Teilnehmer.xlsx', index_col=False)
    teilnehmer = dataframe1.loc[dataframe1['Teilnehmer Nummer'] == int(teilnehmer_nummer)]
    return teilnehmer