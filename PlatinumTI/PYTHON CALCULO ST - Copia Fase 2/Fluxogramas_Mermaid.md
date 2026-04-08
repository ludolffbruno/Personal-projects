# Fluxogramas ST/DIFAL — Mermaid

Cole o bloco abaixo em https://mermaid.live — os dois fluxos aparecem lado a lado.

---

## FASE 1 + FASE 2 — Diagrama Combinado (lado a lado)

```mermaid
flowchart LR

    subgraph F1["📥  FASE 1 — COMPRAS  (Entrada NF)"]
        direction TB
        A1([🟢 INÍCIO])
        A1 --> B1[/"📄 Recebe NF-e\nDigita o NCM\nno aplicativo"/]
        B1 --> C1["🔍 Sistema localiza\ndados do NCM\nna tabela Excel"]
        C1 --> D1{{"Produto sujeito\na ICMS-ST?\nC/ST ou S/ST"}}

        D1 -- "❌ S/ST" --> E1["Selecionar\nAlíquota Origem\n0% / 4% / 12% / 22%"]
        D1 -- "✅ C/ST" --> F1{{"ST já recolhida\nna compra?\nST Inclusa?"}}

        F1 -- "✅ SIM\nST já paga" --> G1["Custo Final =\nCusto Unitário\nsem acréscimo"]
        F1 -- "❌ NÃO" --> E1

        E1 --> H1["💰 CALCULAR\n\nCusto Final =\nCusto Unit.\n+ Custo Unit. × %"]
        H1 --> Z1([🔴 FIM])
        G1 --> Z1
    end

    subgraph F2["📤  FASE 2 — SAÍDAS  (Venda Interestadual | Origem: RJ)"]
        direction TB
        A2([🟢 INÍCIO])
        A2 --> B2[/"📦 Recebe pedido de venda\nSeleciona:\n• UF Destino\n• Importado? Sim/Não\n• NCM / Produto"/]
        B2 --> C2["🔍 Sistema localiza:\n• Acordo ST com a UF\n• MVA  •  Alíq. Inter.\n• Alíq. Interna destino"]
        C2 --> D2{{"Existe Acordo ST\ncom esta UF?"}}

        D2 -- "❌ NÃO" --> E2["Sem ST a reter\nEmite NF normalmente\nICMS próprio apenas"]
        E2 --> Z2([🔴 FIM])

        D2 -- "✅ SIM" --> F2{{"Venda para\nConsumidor Final?"}}

        F2 -- "❌ NÃO\nRevenda" --> G2["Usa MVA\nda planilha"]
        F2 -- "✅ SIM" --> H2{{"Cliente tem IE?\nContribuinte ICMS"}}

        G2 --> MA["📊 MOD. A — REVENDA\nICMS-ST com MVA\n\n1. VT = Valor + Frete\n2. ICMS Prop. = VT × Inter.\n3. Base ST = VT × (1+MVA%)\n4. ICMS-ST = Base ST × Intra\n           − ICMS Prop.\n5. Total NF = VT + ICMS-ST"]

        H2 -- "❌ NÃO\nPessoa Física / sem IE" --> MB["📊 MOD. B — BASE ÚNICA\nCFOP 6108\n\nPreço = valor SEM imposto\n\n1. Novo Valor = Valor ÷ (1−Intra)\n2. ICMS Prop. = Novo Val × Inter.\n3. DIFAL = Novo Val × Intra\n          − ICMS Prop.\n4. Total NF = Novo Valor\n   (imposto embutido)"]

        H2 -- "✅ SIM\nEmpresa com IE" --> MC["📊 MOD. C — BASE DUPLA\nCFOP 6403\n\n1. VT = Valor + Frete\n2. ICMS Prop. = VT × Inter.\n3. Base Dupla = (VT − ICMS Prop.)\n              ÷ (1 − Intra)\n4. DIFAL = Base Dupla × Intra\n          − ICMS Prop.\n5. Total NF = VT + DIFAL"]

        MA --> RES2["� Demonstrativo\nICMS-ST / DIFAL a Reter\nTotal da Nota Fiscal"]
        MB --> RES2
        MC --> RES2
        RES2 --> Z2
    end

    style A1 fill:#217346,color:#fff
    style Z1 fill:#C00000,color:#fff
    style A2 fill:#217346,color:#fff
    style Z2 fill:#C00000,color:#fff
    style E2 fill:#FDECEA
    style MA fill:#E3F2FD
    style MB fill:#F3E5F5
    style MC fill:#E8F5E9
    style G1 fill:#E8F5E9
    style H1 fill:#E3F2FD
    style D1 fill:#FFF8E1
    style F1 fill:#FFF8E1
    style D2 fill:#FFF8E1
    style F2 fill:#FFF8E1
    style H2 fill:#FFF8E1
    style F1 fill:#EEF4FF,stroke:#1F3864,stroke-width:2px,color:#000
    style F2 fill:#F0FFF4,stroke:#217346,stroke-width:2px,color:#000
```
