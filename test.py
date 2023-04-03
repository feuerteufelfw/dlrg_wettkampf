import pandas as pd

df = pd.read_csv('files/teilnehmer.csv',
                 sep='[;]',
                 engine='python')

# Print the Dataframe
print(df)