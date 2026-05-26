import sqlite3
from decimal import Decimal
from typing import Optional
from dataclasses import dataclass

@dataclass
class ResolvedContract:
    listino_r: Optional[Decimal]
    sconto_1: Decimal = Decimal("0.00")
    sconto_2: Decimal = Decimal("0.00")
    sconto_3: Decimal = Decimal("0.00")
    sconto_4: Decimal = Decimal("0.00")
    sconto_5: Decimal = Decimal("0.00")
    sconto_6: Decimal = Decimal("0.00")
    sconto_7: Decimal = Decimal("0.00")
    sconto_carico: Decimal = Decimal("0.00")
    sconto_pagamento: Decimal = Decimal("0.00")
    voce_i: Decimal = Decimal("0.00")
    voce_ii: Decimal = Decimal("0.00")
    voce_iii: Decimal = Decimal("0.00")
    voce_iv: Decimal = Decimal("0.00")
    voce_v: Decimal = Decimal("0.00")
    livello_risolto: str = "NESSUNO"

class HierarchyResolver:
    _FIELDS = {
        "sconto_1": "sconto_1", "sconto_2": "sconto_2", "sconto_3": "sconto_3",
        "sconto_4": "sconto_4", "sconto_5": "sconto_5", "sconto_6": "sconto_6",
        "sconto_7": "sconto_7", "sconto_carico": "sconto_carico", 
        "sconto_pagamento": "sconto_pagamento", "voce_contratto_1": "voce_i",
        "voce_contratto_2": "voce_ii", "voce_contratto_3": "voce_iii",
        "voce_contratto_4": "voce_iv", "voce_contratto_5": "voce_v"
    }

    @classmethod
    def resolve(cls, conn: sqlite3.Connection, gruppo: str, sottogruppo: str, insegna: str, ean: str, categoria: str) -> ResolvedContract:
        conn.row_factory = sqlite3.Row
        cursor = conn.cursor()

        cursor.execute("""
            SELECT livello, chiave_livello, listino_r,
                   sconto_1, sconto_2, sconto_3, sconto_4, sconto_5,
                   sconto_6, sconto_7, sconto_carico, sconto_pagamento,
                   voce_contratto_1, voce_contratto_2, voce_contratto_3, 
                   voce_contratto_4, voce_contratto_5
            FROM accordi_commerciali
            WHERE (UPPER(TRIM(gruppo_macro)) = UPPER(TRIM(?)) AND livello = 'GRUPPO' AND (sottogruppo = '' OR sottogruppo IS NULL) AND (associato_insegna = '' OR associato_insegna IS NULL))
               OR (UPPER(TRIM(gruppo_macro)) = UPPER(TRIM(?)) AND UPPER(TRIM(sottogruppo)) = UPPER(TRIM(?)) AND livello = 'SOTTOGRUPPO' AND (associato_insegna = '' OR associato_insegna IS NULL))
               OR (UPPER(TRIM(gruppo_macro)) = UPPER(TRIM(?)) AND UPPER(TRIM(sottogruppo)) = UPPER(TRIM(?)) AND UPPER(TRIM(associato_insegna)) = UPPER(TRIM(?)) AND livello = 'CATEGORIA' AND UPPER(TRIM(chiave_livello)) = UPPER(TRIM(?)))
               OR (UPPER(TRIM(gruppo_macro)) = UPPER(TRIM(?)) AND UPPER(TRIM(sottogruppo)) = UPPER(TRIM(?)) AND UPPER(TRIM(associato_insegna)) = UPPER(TRIM(?)) AND livello = 'REFERENZA' AND chiave_livello = ?)
            ORDER BY 
                CASE livello
                    WHEN 'GRUPPO' THEN 1
                    WHEN 'SOTTOGRUPPO' THEN 2
                    WHEN 'CATEGORIA' THEN 3
                    WHEN 'REFERENZA' THEN 4
                END ASC
        """, (gruppo, gruppo, sottogruppo, gruppo, sottogruppo, insegna, categoria, gruppo, sottogruppo, insegna, ean))

        rules = cursor.fetchall()
        contract = ResolvedContract(listino_r=None)

        for row in rules:
            liv = row["livello"]
            chiave = (row["chiave_livello"] or "").upper().strip()

            if liv == "CATEGORIA" and chiave != categoria.upper().strip():
                continue
            if liv == "REFERENZA" and chiave != ean.strip():
                continue

            if row["listino_r"] is not None:
                contract.listino_r = Decimal(str(row["listino_r"]))

            for db_field, attr in cls._FIELDS.items():
                val = row[db_field]
                if val is not None:
                    setattr(contract, attr, Decimal(str(val)))
            
            contract.livello_risolto = liv

        return contract
