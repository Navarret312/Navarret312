from azure.storage.blob import BlobServiceClient

# Nome da conta de armazenamento e chave de acesso
storage_account_name = "sqlblobstoragedatabricks"
storage_account_access_key = ""
container_name = "output"

# Criar a URL do serviço de blob
blob_service_client = BlobServiceClient(
    account_url=f"https://{storage_account_name}.blob.core.windows.net",
    credential=storage_account_access_key
)

# Nome do "diretório"
directory_name = "servicenow/"

# Criar um blob vazio para representar a pasta
blob_client = blob_service_client.get_blob_client(container=container_name, blob=directory_name + "placeholder.txt")
blob_client.upload_blob(b"", overwrite=True)

print(f"Pasta '{directory_name}' criada com sucesso no container '{container_name}'!")