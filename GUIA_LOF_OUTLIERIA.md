# Guia simples: Outlieria com LOF espacial

## 1. O que foi adicionado

O Outlieria mantém o IQR como análise estatística e adiciona uma segunda avaliação opcional chamada **LOF espacial**. O LOF usa somente o histórico anterior a 2026 e compara cada ponto com vizinhos da mesma lavoura e profundidade, dentro do raio informado na aba Fontes.

O resultado final pode ser:

- **outlier confirmado**: IQR e LOF identificaram anomalia;
- **outlier estatístico**: somente o IQR identificou anomalia;
- **anomalia espacial**: somente o LOF identificou anomalia;
- **atenção**: ponto próximo do limite ou com dados insuficientes;
- **normal**: nenhuma anomalia identificada.

## 2. Preparar o ambiente

No terminal, entre na pasta do projeto:

```bash
cd /caminho/para/outlieria
```

Crie um ambiente virtual:

```bash
python -m venv .venv
```

Ative-o no Linux/macOS:

```bash
source .venv/bin/activate
```

Ative-o no Windows PowerShell:

```powershell
.venv\\Scripts\\Activate.ps1
```

Instale as dependências:

```bash
python -m pip install --upgrade pip
pip install -r requirements.txt
```

O pacote novo mais importante é:

```text
scikit-learn
```

## 3. Executar a aplicação

Com o ambiente virtual ativo:

```bash
streamlit run streamlit_app.py
```

Abra o endereço mostrado pelo Streamlit, normalmente:

```text
http://localhost:8501
```

## 4. Usar o LOF na interface

1. Abra a aba **Fontes**.
2. Envie a planilha Excel.
3. Envie o limite SHP.
4. Informe o raio regional em metros.
5. Use o raio do maior grid quando os grids tiverem tamanhos diferentes.
6. Marque **Ativar LOF espacial**.
7. Clique em **Pronto**.

O processamento seguirá esta ordem:

```text
Excel + SHP
→ validação do raio
→ espacialização
→ IQR regional
→ LOF espacial
→ combinação IQR + LOF
→ interface e exportação
```

## 5. Requisitos mínimos para o LOF

O modelo precisa de histórico suficiente. Por grupo de lavoura, profundidade e elemento, ele tenta usar pelo menos 10 registros históricos treináveis e pelo menos 5 vizinhos por ponto.

Quando não houver dados suficientes, o sistema não inventa um score. Ele registra **dados insuficientes** e mantém o resultado estatístico disponível.

## 6. O que o LOF usa

Para cada ponto, o sistema calcula:

- valor do elemento;
- média e mediana regionais;
- Q1, Q3 e IQR regionais;
- desvio-padrão regional;
- quantidade de vizinhos;
- distância média dos vizinhos;
- distância até o vizinho mais próximo;
- diferença para a mediana;
- z-score regional.

Os elementos ignorados pela regra do Outlieria continuam fora do LOF.

## 7. Conferir o resultado

Na tabela e no tooltip do gráfico, podem aparecer:

- resultado do IQR;
- resultado LOF;
- score LOF;
- quantidade de vizinhos;
- classificação final.

As cores representam:

- vermelho: outlier confirmado;
- laranja: outlier estatístico;
- roxo: anomalia espacial;
- amarelo: atenção;
- azul-esverdeado: normal.

## 8. Publicar no Streamlit Cloud

O `requirements.txt` já contém o scikit-learn. Ao enviar o projeto ao GitHub, o Streamlit Cloud instala a dependência automaticamente.

Se a aplicação já estiver publicada:

1. abra o repositório no GitHub;
2. confirme que `scikit-learn>=1.5,<2` está no `requirements.txt`;
3. aguarde o redeploy automático;
4. use **Manage app → Reboot app** se necessário;
5. teste primeiro com o LOF desativado;
6. depois marque **Ativar LOF espacial** em uma análise com histórico suficiente.

## 9. Cuidados de interpretação

O LOF é uma segunda opinião estatística, não uma confirmação agronômica automática. Um ponto anômalo pode representar erro de coleta, erro de digitação, mudança real do solo ou uma transição espacial legítima.

Por isso, a recomendação é começar com o LOF ativado em análises de teste e comparar os resultados com o IQR e com a avaliação técnica de campo.
