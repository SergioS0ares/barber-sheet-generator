# 💈 Barber Sheet Generator

> Automação de planilhas financeiras para Barbearias usando Python.

Este projeto consiste em um script Python que automatiza a criação de arquivos Excel (`.xlsx`) detalhados para o controle financeiro mensal de uma barbearia. O script gera uma planilha pronta para uso, com abas para todos os dias do mês, fórmulas automáticas e visual profissional.

## ✨ Funcionalidades

* **Geração Dinâmica:** Cria abas para todos os dias do mês automaticamente, respeitando anos bissextos e dias totais (28, 29, 30 ou 31).
* **Design Profissional:**
    * Estilo "Zebrado" (linhas alternadas em azul e branco) para facilitar a leitura.
    * Bordas formatadas e cabeçalhos destacados.
    * Painéis congelados (Freeze Panes) para manter o cabeçalho visível ao rolar.
* **Automação de Fórmulas:**
    * Cálculo automático de lucro (Venda - Custo) para itens secundários (Picolé, Bebidas).
    * Somatórios automáticos no rodapé de cada dia.
* **Resumo Mensal:** Uma aba final "TOTAL DO MÊS" que consolida os dados de todas as abas diárias em um relatório financeiro completo.
* **Validação de Dados:** Listas suspensas (Dropdowns) para seleção de forma de pagamento (PIX, Dinheiro, Cartão), evitando erros de digitação.

## 🛠️ Tecnologias Utilizadas

* [Python](https://www.python.org/)
* [Pandas](https://pandas.pydata.org/) (Manipulação de dados)
* [XlsxWriter](https://xlsxwriter.readthedocs.io/) (Motor de geração do Excel e formatação condicional)

## 🚀 Como usar

### Pré-requisitos

Você precisa ter o Python instalado. Em seguida, instale as bibliotecas necessárias:

```bash
pip install pandas xlsxwriter
