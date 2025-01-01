import pandas as pd

data = {
    'Usuario': [
        "JAIME GALVAN BALBOA",
        "JUAN PEREZ",
        "MARIA GONZALEZ RAMIREZ",
        "ANA",
        None
    ]
}

semaforo_df = pd.DataFrame(data)

semaforo_df['Usuario_Semaforo'] = semaforo_df['Usuario'].apply(
    lambda x: ' '.join([x.split()[0], x.split()[2]]) if isinstance(x, str) and len(x.split()) == 4 else
              ' '.join([x.split()[0], x.split()[1]]) if isinstance(x, str) and len(x.split()) == 3 else
              x if isinstance(x, str) else ''
)
print(semaforo_df)
