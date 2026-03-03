# 02 Model Flowchart

```mermaid
flowchart LR
    subgraph Other
        _[ ]
        QCF[QCF]
        Cons[Cons]
        Ops[Ops]
        LandN[L&N]
    end
    subgraph Inputs_/_Assumptions
        Inputs[Inputs]
    end
    subgraph Calculations
        Debt[Debt]
    end
    subgraph Timing
        Timing[Timing]
    end
    Timing --> Ops
    Timing --> Debt
    Inputs --> Cons
    Timing --> QCF
    Cons --> Debt
    Ops --> QCF
    Inputs --> Debt
    Debt --> QCF
    Inputs --> Ops
    QCF --> Debt
    QCF --> Ops
    Inputs --> Timing
    Timing --> Cons
    Cons --> Ops
    Cons --> QCF
```