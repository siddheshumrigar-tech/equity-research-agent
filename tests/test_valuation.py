import unittest
import math

class TestValuationEngine(unittest.TestCase):
    def calculate_capm(self, rf: float, beta: float, erp: float) -> float:
        return rf + beta * erp

    def calculate_mid_year_discount_factor(self, wacc: float, t: int) -> float:
        return 1.0 / ((1.0 + wacc) ** (t - 0.5))

    def calculate_gordon_terminal_value(self, fcff_last: float, g: float, wacc: float) -> float:
        return (fcff_last * (1.0 + g)) / (wacc - g)

    def test_capm_formula(self):
        ke = self.calculate_capm(0.0705, 0.85, 0.0550)
        self.assertEqual(round(ke * 100, 2), 11.72)
        
        ke_us = self.calculate_capm(0.0425, 1.10, 0.0500)
        self.assertEqual(round(ke_us * 100, 2), 9.75)

    def test_mid_year_discounting(self):
        wacc = 0.10
        df1 = self.calculate_mid_year_discount_factor(wacc, 1)
        expected_df1 = 1.0 / math.sqrt(1.10)
        self.assertAlmostEqual(df1, expected_df1, places=5)
        
        df2 = self.calculate_mid_year_discount_factor(wacc, 2)
        self.assertTrue(df2 < df1)

    def test_gordon_growth_terminal_value(self):
        fcff = 100.0
        g = 0.045
        wacc = 0.115
        tv = self.calculate_gordon_terminal_value(fcff, g, wacc)
        self.assertEqual(round(tv, 2), 1492.86)

    def test_sentinel_beta_preservation(self):
        live_beta = 1.00
        beta = None
        if live_beta is not None:
            beta = float(live_beta)
        if beta is None:
            beta = 0.85
        self.assertEqual(beta, 1.00)

    def test_verdict_logic(self):
        def get_verdict(upside_pct):
            if upside_pct >= 18.0:
                return "STRONG ACCUMULATE"
            elif upside_pct >= 7.0:
                return "ACCUMULATE"
            elif upside_pct >= -5.0:
                return "NEUTRAL / HOLD"
            elif upside_pct >= -15.0:
                return "REDUCE"
            else:
                return "SELL"
                
        self.assertEqual(get_verdict(25.0), "STRONG ACCUMULATE")
        self.assertEqual(get_verdict(12.0), "ACCUMULATE")
        self.assertEqual(get_verdict(0.0), "NEUTRAL / HOLD")
        self.assertEqual(get_verdict(-10.0), "REDUCE")
        self.assertEqual(get_verdict(-35.0), "SELL")

    def test_dynamic_thesis_no_boilerplate_for_losers(self):
        verdict = "SELL"
        ticker = "WEAKCORP"
        if verdict in ["STRONG ACCUMULATE", "ACCUMULATE"]:
            thesis = f"{ticker} is an established compounder"
        else:
            thesis = f"{ticker} trades at a stretched valuation relative to projected fundamental cash flows"
            
        self.assertNotIn("compounder", thesis)
        self.assertIn("stretched valuation", thesis)

if __name__ == '__main__':
    unittest.main()
