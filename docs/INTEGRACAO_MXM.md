# Integração GCom → MXM (CasaTua)

Plano para enviar os títulos a pagar extraídos do GCom direto para o MXM da CasaTua,
no mesmo modelo usado pelo projeto RioEnergy.TUST.

## Como o TUST integra com o MXM

Referência: `RioEnergy.TUST.ClassesNegocio/Integracao/MXMIntegracaoTituloPagar.cs`.

O TUST não usa planilha nem acessa o banco do MXM: chama a **API REST de interface do
Contas a Pagar** do MXM, em duas etapas assíncronas.

### 1. Gravar títulos

`POST {URLBaseAPIMXM}/InterfacedoContasPagarReceber/Gravar`

```json
{
  "AutheticationToken": { "Username": "...", "Password": "...", "EnvironmentName": "..." },
  "Data": {
    "InterfacedoContasPagarReceber": [
      {
        "SequenciadoRegistro": 1,
        "Identificacao": "PP",
        "...": "...",
        "InterfaceGrupoPagarReceber": [
          { "SequenciadoGrupoPagarReceber": "1", "CodigodoGrupo": "...", "ValordoGrupo": "680.76" }
        ]
      }
    ]
  }
}
```

- `AutheticationToken` está escrito errado assim mesmo: é o nome que a API espera.
- A resposta traz `Success` e uma mensagem contendo **"processo N"**; esse número é guardado.

### 2. Consultar o resultado

`POST {URLBaseAPIMXM}/InterfacedoContasPagarReceber/ConsultaporProcesso`
com `Data: { "SequenciadoProcesso": N }`.

- No TUST, o robô `ConsultarRetornoMXM` consulta a cada 5 minutos.
- Erros por título vêm em `InterfaceErrodoContasPagarReceber[].CodigodoErro` / `MensagemdoErro`.
- "Registro não processado" = ainda na fila, não é erro.

Credenciais ficam na tabela de parâmetros do TUST: `URLBaseAPIMXM`, `UsernameMXM`,
`PasswordMXM`, `EnvironmentNameMXM`.

## O template GCom já é o layout do MXM

`template/TEMPLATE_TITULOS_A_PAGAR_(GCOM).xlsx` é o layout de importação de títulos do MXM.
As colunas mapeiam 1:1 para os campos da API, então dá para pular a planilha e enviar JSON:

| Coluna do template | Valor atual | Campo na API |
|---|---|---|
| A Identificação | `PP` | `Identificacao` |
| B Codigo do Fornecedor | CNPJ (dígitos) | `CodigoClienteFornecedor` |
| C Numero do Titulo | `108839 - 1 / 1` | `NumerodoTitulo` |
| D Documento Fiscal | idem | `DocumentoFiscal` |
| E Empresa Emitente / G Empresa Pagadora | `0001` | `EmpresaEmitente` / `EmpresaRecebedora` |
| F Codigo da Filial | `0001` | `Filial` |
| H Tipo de Titulo | `55` | `TipodeTitulo` |
| I / J / K Emissão / Vencimento / Programação | `ddMMyyyy` | `DatadeEmissao` / `DatadeVencimento` / `DatadaProgramacao` |
| L Moeda / M Tipo de Cobrança | `BRL` / `CA` | `CodigodaMoeda` / `TipodeCobranca` |
| CE Grupo de Pagamento | `1106010000` | `InterfaceGrupoPagarReceber[].CodigodoGrupo` |
| CG Valor do Grupo de Pagamento | `680.76` | `ValordoGrupo` e `ValordoTitulo` (2 casas, ponto decimal) |
| CJ Codigo do fluxo de caixa | `01` | `InterfaceGrupoPagarReceber[].ContadoFluxodeCaixa` |

Campos que o TUST preenche e o template não tem: `DatadaEntrada`, `DatadeCompetencia`,
`FormadePagamento`, `CodigodaContaDocumento`, centro de custo.

## O que precisamos do MXM da CasaTua

### A. Acesso (equipe/suporte MXM da CasaTua)

1. **URL base da API REST** da instância CasaTua.
2. **Usuário de integração** (dedicado, não pessoal) com permissão em *Interface do Contas a Pagar/Receber*.
3. **EnvironmentName** — **homologação primeiro**, produção depois.
4. Rede: a API é acessível de onde o app roda ou exige VPN / liberação de IP?

### B. Cadastros e códigos (contabilidade/financeiro)

5. **Código do fornecedor no MXM:** é o CNPJ ou código interno? Se interno, precisamos de
   consulta ou de-para. Todo fornecedor do GCom precisa existir no MXM.
6. **Empresa / Filial / Empresa Pagadora** (hoje fixos `0001`) e **Tipo de Título** (`55`).
7. **Grupo de Pagamento** (`1106010000`), **conta de fluxo de caixa** (`01`) e se
   **centro de custo** é obrigatório.
8. **Tipo de cobrança** (`CA`) e **forma de pagamento** (obrigatória na API, ausente no template).
9. **Formato de `NumerodoTitulo`:** o MXM aceita `108839 - 1 / 1` (tamanho máximo, caracteres)
   ou devemos mandar `108839-1/1`? Há checagem de duplicidade por fornecedor + número?
10. **Data de entrada e competência:** usar a "DATA DA ENTRADA" do GCom?

### C. Implementação (após A e B)

11. `services/mxm_client.py` com `gravar(titulos)` e `consultar(processo)`, e um mapeamento
    que gera o JSON da tabela acima a partir do mesmo DataFrame que monta a planilha.
12. Credenciais MXM em `.streamlit/secrets.toml` / variáveis de ambiente (como as do GCom).
13. Fluxo: gerar relatório → prévia → enviar ao MXM → nº do processo → consultar retorno
    (sob demanda ou em loop) → mostrar erros por título.
14. Primeiro teste em homologação com um único título.

**Para começar:** itens 1–3 (acesso de homologação) e a resposta do item 5 (código do fornecedor).
