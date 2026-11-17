# Configuração do armazenamento do Outlieria

A arquitetura usa o Supabase/PostgreSQL para metadados e o Google Drive para os arquivos binários. O aplicativo continua funcionando em modo de sessão quando as Secrets ainda não foram configuradas.

## 1. Supabase

1. Crie um projeto no Supabase.
2. Abra o SQL Editor.
3. Execute o conteúdo de `supabase_schema.sql`.
4. Em **Project Settings → API**, copie a URL do projeto e a chave `service_role`.

A chave `service_role` deve ficar somente nas Secrets do Streamlit Cloud. Ela não deve ser publicada no GitHub nem enviada ao navegador.

## 2. Google Drive

Crie um projeto OAuth no Google Cloud, ative a Google Drive API e gere um refresh token para a conta pessoal que será usada pelo Outlieria. A aplicação usa o escopo restrito `drive.file`, suficiente para criar e atualizar os arquivos gerenciados pelo app.

O refresh token é obtido uma única vez em um fluxo OAuth autorizado pelo usuário. Não coloque `client_secret.json` no repositório.

## 3. Streamlit Cloud

Em **Settings → Secrets**, adicione:

```toml
SUPABASE_URL = "https://SEU-PROJETO.supabase.co"
SUPABASE_SERVICE_ROLE_KEY = "sua-chave-service-role"
GOOGLE_CLIENT_ID = "seu-client-id.apps.googleusercontent.com"
GOOGLE_CLIENT_SECRET = "seu-client-secret"
GOOGLE_REFRESH_TOKEN = "seu-refresh-token"
```

Depois de salvar as Secrets, reinicie a aplicação.

## 4. Comportamento no aplicativo

Ao clicar em **Pronto**, o Outlieria:

- processa o Excel e o limite SHP;
- calcula os outliers e a reanálise de 2026;
- grava Excel original, SHP e reanálise nas pastas do Drive;
- grava cliente, fazenda, área, quantidade de pontos, contagens e hash no Supabase;
- grava também o raio regional usado naquela análise;
- reutiliza os mesmos arquivos do Drive quando a mesma combinação de Excel + SHP é enviada novamente.

A estrutura criada no Drive é:

```text
Outlieria/
├── 01_Fontes_Originais/
└── 02_Reanalises_Individuais/
```

Sem as cinco Secrets, nenhuma chamada externa é feita e o processamento permanece apenas na sessão, permitindo testar a interface com segurança.

## 5. Raio regional

Na aba **Fontes**, informe o raio em metros no campo **Raio regional (m)**. O valor padrão é 300 m, mas pode ser alterado para cada novo grid. O cálculo usa esse valor para selecionar o histórico de cada ponto; a reanálise exportada também registra o raio nos comentários dos outliers.
