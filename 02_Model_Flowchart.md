# 02 Model Flowchart

```mermaid
flowchart LR
    subgraph Cover_/_TOC
        Cover[Cover]
    end
    subgraph Control_/_Scenario
        Control[Control]
    end
    subgraph Timing
        Timeline[Timeline]
    end
    subgraph Other
        SKU_Matrix[SKU_Matrix]
        Assum_Volume[Assum_Volume]
        Assum_Pricing[Assum_Pricing]
        Assum_Costs[Assum_Costs]
        Assum_WC[Assum_WC]
        Assum_Finance[Assum_Finance]
        CF[CF]
    end
    subgraph Calculations
        Assum_Capex[Assum_Capex]
        Calc_Volume[Calc_Volume]
        Calc_Revenue[Calc_Revenue]
        Calc_COGS[Calc_COGS]
        Calc_Opex[Calc_Opex]
        Calc_Capex[Calc_Capex]
        Calc_WC[Calc_WC]
        Calc_Debt[Calc_Debt]
    end
    subgraph Financial_Statements
        IS[IS]
        BS[BS]
        KPIs[KPIs]
    end
    subgraph Checks
        Checks[Checks]
    end
    subgraph Outputs_/_Dashboard
        Annual_Summary[Annual_Summary]
    end
    Assum_Costs --> Calc_COGS
    Assum_WC --> KPIs
    Calc_Volume --> Calc_Opex
    Calc_Volume --> BS
    BS --> Annual_Summary
    IS --> KPIs
    Calc_Volume --> Calc_Debt
    CF --> Annual_Summary
    BS --> Checks
    Assum_Costs --> Calc_Opex
    CF --> Checks
    Calc_Volume --> Calc_Capex
    Assum_Capex --> Calc_Capex
    Calc_COGS --> IS
    IS --> BS
    Calc_Capex --> Annual_Summary
    Assum_Pricing --> Calc_Revenue
    Assum_Finance --> CF
    Assum_Finance --> Checks
    Calc_Capex --> CF
    Calc_COGS --> Calc_Opex
    Control --> Calc_COGS
    Calc_Revenue --> Annual_Summary
    Control --> IS
    CF --> KPIs
    BS --> KPIs
    Calc_WC --> CF
    Calc_Volume --> Calc_Revenue
    Calc_Debt --> Annual_Summary
    Calc_Revenue --> Calc_WC
    Calc_Debt --> CF
    Calc_Debt --> Checks
    Assum_Volume --> Calc_Volume
    Calc_Volume --> Annual_Summary
    Calc_Capex --> KPIs
    CF --> BS
    Calc_Capex --> IS
    Calc_Volume --> CF
    Calc_Volume --> Checks
    Calc_Opex --> Calc_WC
    Calc_Volume --> Calc_WC
    Control --> Calc_Volume
    Assum_Volume --> Calc_Revenue
    Calc_Revenue --> KPIs
    Calc_WC --> KPIs
    Calc_Revenue --> Calc_COGS
    Calc_Revenue --> IS
    Calc_Capex --> BS
    Calc_Debt --> KPIs
    IS --> Annual_Summary
    Assum_Finance --> BS
    Assum_Volume --> Assum_WC
    Calc_Debt --> IS
    IS --> CF
    IS --> Checks
    Assum_Finance --> Calc_Debt
    Assum_WC --> Calc_WC
    Control --> Calc_Revenue
    Calc_Volume --> KPIs
    Calc_WC --> BS
    Calc_Revenue --> Calc_Opex
    Calc_Volume --> Calc_COGS
    Assum_Volume --> Checks
    Calc_Opex --> IS
    Calc_Volume --> IS
    Calc_Debt --> BS
    Calc_COGS --> Annual_Summary
    Calc_COGS --> Calc_WC
```