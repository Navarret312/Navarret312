import speech_recognition as sr

# Cria um reconhecedor
recognizer = sr.Recognizer()

# Carrega o arquivo de áudio
with sr.AudioFile("output.wav") as source:
    print("Reconhecendo o áudio...")
    audio_data = recognizer.record(source)

# Tenta reconhecer o áudio usando o reconhecimento de fala do Google
try:
    texto = recognizer.recognize_google(audio_data, language="pt-BR")
    print("Texto reconhecido:")
    print(texto)
except sr.UnknownValueError:
    print("Não foi possível entender o áudio.")
except sr.RequestError as e:
    print(f"Erro ao se conectar ao serviço de reconhecimento: {e}")
