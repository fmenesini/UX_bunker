from decimal import Decimal, ROUND_HALF_UP
from typing import List
from dataclasses import dataclass

@dataclass(frozen=True)
class PricingInput:
    listino_r: Decimal
    sconto_1: Decimal = Decimal("0.00")
    sconto_2: Decimal = Decimal("0.00")
    sconto_3: Decimal = Decimal("0.00")
    sconto_4: Decimal = Decimal("0.00")
    sconto_5: Decimal = Decimal("0.00")
    sconto_6: Decimal = Decimal("0.00")
    sconto_7: Decimal = Decimal("0.00")
    sconto_y: Decimal = Decimal("0.00")
    sconto_z: Decimal = Decimal("0.00")
    sconto_aa: Decimal = Decimal("0.00")
    sconto_carico: Decimal = Decimal("0.00")
    sconto_pagamento: Decimal = Decimal("0.00")
    voce_i: Decimal = Decimal("0.00")
    voce_ii: Decimal = Decimal("0.00")
    voce_iii: Decimal = Decimal("0.00")
    voce_iv: Decimal = Decimal("0.00")
    voce_v: Decimal = Decimal("0.00")
    min_net_net_g: Decimal = Decimal("0.00")

@dataclass(frozen=True)
class WaterfallStep:
    fase: str
    valore: Decimal
    descrizione: str

@dataclass
class PricingResult:
    steps: List[WaterfallStep]
    netto_in_fattura_2: Decimal
    contratto_tot_pfa: Decimal
    net_net_finale: Decimal
    delta_vs_min: Decimal
    guardrail_ok: bool
    sconto_max_av: Decimal

class PricingEngine:
    @staticmethod
    def _apply_pct(base: Decimal, pct: Decimal) -> Decimal:
        if pct < 0 or pct >= 100:
            raise ValueError(f"Percentuale sconto fuori limite consentito: {pct}%")
        factor = (Decimal("100.00") - pct) / Decimal("100.00")
        return base * factor

    @classmethod
    def calculate(cls, inp: PricingInput) -> PricingResult:
        steps: List[WaterfallStep] = []
        prezzo = inp.listino_r
        
        steps.append(WaterfallStep("1. Listino Base (R)", prezzo, "Prezzo di partenza contrattuale"))

        centrali = [inp.sconto_1, inp.sconto_2, inp.sconto_3, inp.sconto_4, inp.sconto_5]
        sconti_attivi_c = [f"{s}%" for s in centrali if s > 0]
        for s in centrali:
            if s > 0:
                prezzo = cls._apply_pct(prezzo, s)
        steps.append(WaterfallStep("2. Sconti Centrali (S1-S5)", prezzo, " x ".join(sconti_attivi_c) if sconti_attivi_c else "Nessuno"))

        locali = [inp.sconto_6, inp.sconto_7]
        sconti_attivi_l = [f"{s}%" for s in locali if s > 0]
        for s in locali:
            if s > 0:
                prezzo = cls._apply_pct(prezzo, s)
        steps.append(WaterfallStep("3. Sconti Locali (S6-S7)", prezzo, " x ".join(sconti_attivi_l) if sconti_attivi_l else "Nessuno"))

        if inp.sconto_y > 0:
            prezzo = cls._apply_pct(prezzo, inp.sconto_y)
            steps.append(WaterfallStep("4. Sconto Continuativo (Y)", prezzo, f"-{inp.sconto_y}%"))
        else:
            steps.append(WaterfallStep("4. Sconto Continuativo (Y)", prezzo, "Non applicato"))

        if inp.sconto_z > 0:
            prezzo = cls._apply_pct(prezzo, inp.sconto_z)
            steps.append(WaterfallStep("5. Sconto Promozionale (Z)", prezzo, f"-{inp.sconto_z}%"))
        else:
            steps.append(WaterfallStep("5. Sconto Promozionale (Z)", prezzo, "Non applicato"))

        if inp.sconto_aa > 0:
            prezzo = max(Decimal("0.00"), prezzo - inp.sconto_aa)
            steps.append(WaterfallStep("6. Sconto Unitario in fattura (AA)", prezzo, f"-{inp.sconto_aa:.2f} Euro/Pz"))
        else:
            steps.append(WaterfallStep("6. Sconto Unitario in fattura (AA)", prezzo, "Non applicato"))

        if inp.sconto_carico > 0:
            prezzo = cls._apply_pct(prezzo, inp.sconto_carico)
        if inp.sconto_pagamento > 0:
            prezzo = cls._apply_pct(prezzo, inp.sconto_pagamento)
        
        netto_fatt_2 = prezzo.quantize(Decimal("0.001"), rounding=ROUND_HALF_UP)
        steps.append(WaterfallStep("7. Netto in Fattura 2 (AF)", netto_fatt_2, f"Logistica (AB): -{inp.sconto_carico}% | Finanziario (AC): -{inp.sconto_pagamento}%"))

        pfa_sum = inp.voce_i + inp.voce_ii + inp.voce_iii + inp.voce_iv + inp.voce_v
        if pfa_sum >= 100:
            raise ValueError(f"La somma dei Premi Fuori Fattura ({pfa_sum}%) supera o eguaglia il 100%. Ricavo impossibile.")
            
        net_net = cls._apply_pct(netto_fatt_2, pfa_sum)
        net_net = net_net.quantize(Decimal("0.001"), rounding=ROUND_HALF_UP)
        steps.append(WaterfallStep("8. Net Net Finale (AM)", net_net, f"Totale Premi (AL): -{pfa_sum}% (Voci I-V)"))

        delta = net_net - inp.min_net_net_g
        guardrail_ok = delta >= 0

        sconto_max_av = Decimal("0.00")
        if inp.listino_r > 0:
            pfa_factor = (Decimal("100.00") - pfa_sum) / Decimal("100.00")
            log_factor = (Decimal("100.00") - inp.sconto_carico) / Decimal("100.00")
            fin_factor = (Decimal("100.00") - inp.sconto_pagamento) / Decimal("100.00")
            combined = pfa_factor * log_factor * fin_factor
            if combined > 0:
                prezzo_min_necessario = inp.min_net_net_g / combined
                prezzo_pre_aa_minimo = prezzo_min_necessario + inp.sconto_aa
                sconto_max_av = (Decimal("1.00") - (prezzo_pre_aa_minimo / inp.listino_r)) * Decimal("100.00")
                sconto_max_av = max(Decimal("0.00"), sconto_max_av).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)

        return PricingResult(
            steps=steps,
            netto_in_fattura_2=netto_fatt_2,
            contratto_tot_pfa=pfa_sum,
            net_net_finale=net_net,
            delta_vs_min=delta,
            guardrail_ok=guardrail_ok,
            sconto_max_av=sconto_max_av
        )

    @classmethod
    def calculate_inverse(cls, target_net_net: Decimal, inp: PricingInput, target_field: str = "Z") -> Decimal:
        pfa_sum = inp.voce_i + inp.voce_ii + inp.voce_iii + inp.voce_iv + inp.voce_v
        if pfa_sum >= 100:
            return Decimal("0.00")
            
        pfa_factor = (Decimal("100.00") - pfa_sum) / Decimal("100.00")
        netto_fatt_2_req = target_net_net / pfa_factor
        
        log_factor = (Decimal("100.00") - inp.sconto_carico) / Decimal("100.00")
        fin_factor = (Decimal("100.00") - inp.sconto_pagamento) / Decimal("100.00")
        log_fin_combined = log_factor * fin_factor
        
        if log_fin_combined <= 0:
            return Decimal("0.00")
            
        prezzo_ante_log_req = netto_fatt_2_req / log_fin_combined
        prezzo_post_z_req = prezzo_ante_log_req + inp.sconto_aa

        prezzo_base = inp.listino_r
        for s in [inp.sconto_1, inp.sconto_2, inp.sconto_3, inp.sconto_4, inp.sconto_5, inp.sconto_6, inp.sconto_7, inp.sconto_y]:
            if s > 0:
                prezzo_base = cls._apply_pct(prezzo_base, s)

        if target_field == "Z":
            if prezzo_base <= 0:
                return Decimal("0.00")
            z_req = (Decimal("1.00") - (prezzo_post_z_req / prezzo_base)) * Decimal("100.00")
            return max(Decimal("0.00"), z_req).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)

        elif target_field == "AA":
            prezzo_post_z = prezzo_base
            if inp.sconto_z > 0:
                prezzo_post_z = cls._apply_pct(prezzo_post_z, inp.sconto_z)
            aa_req = prezzo_post_z - prezzo_ante_log_req
            return max(Decimal("0.00"), aa_req).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)

        return Decimal("0.00")
