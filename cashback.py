"""
Cashback rules and calculation for each card.
Each function takes (amount, category, description) and returns the cashback amount.
"""

# Cashback rate config per card: list of (category_keyword, rate_as_decimal)
# Categories are checked as case-insensitive substrings.
# 'default' is the fallback rate.

CASHBACK_RULES = {
    "sbi_cb": {  # SBI Cashback — 5% on online spends
        "Groceries": 0.05,
        "Electricity Bill": 0.05,
        "Recharge / Mobile": 0.05,
        "Shopping": 0.05,
        "Dining / Food": 0.05,
        "Entertainment": 0.05,
        "Travel": 0.05,
        "Medical / Health": 0.05,
        "Fuel": 0.01,        # 1% on fuel
        "Rent": 0.0,
        "Insurance": 0.0,
        "EMI / Loan": 0.0,
        "default": 0.05,     # 5% on everything else online
    },
    "sbi_pp": {  # SBI PhonePe — flat 2%
        "default": 0.02,
    },
    "hdfc_mil": {  # HDFC Millennia — 5% on specific categories, 1% rest
        "Groceries": 0.05,
        "Dining / Food": 0.05,
        "Recharge / Mobile": 0.05,
        "Shopping": 0.05,
        "Entertainment": 0.05,
        "Electricity Bill": 0.01,
        "Fuel": 0.01,
        "Travel": 0.01,
        "Medical / Health": 0.01,
        "Rent": 0.0,
        "EMI / Loan": 0.0,
        "default": 0.01,
    },
    "hdfc_swig": {  # HDFC Swiggy — 10% on food/dining, 5% on others
        "Dining / Food": 0.10,
        "Groceries": 0.05,
        "Restaurant": 0.10,
        "Swiggy": 0.10,
        "Zomato": 0.10,
        "default": 0.05,
    },
    "bob_eterna": {  # BOB Eterna — 2.5% on most, 0% on rent/insurance
        "Rent": 0.0,
        "Insurance": 0.0,
        "EMI / Loan": 0.0,
        "Fuel": 0.01,
        "default": 0.025,
    },
}


def calculate_cashback(amount, category, card_id, description=""):
    """
    Calculate expected cashback for a transaction based on card rules.
    Returns (estimated_cashback, rule_label).
    """
    if not card_id or card_id in ("other", "other_card", ""):
        return 0.0, "No card selected"

    rules = CASHBACK_RULES.get(card_id)
    if not rules:
        return 0.0, "Unknown card"

    category_lower = (category or "").lower()
    desc_lower = (description or "").lower()

    # Try exact category match first, then substring match
    rate = rules.get("default", 0.0)
    matched_rule = "default"

    for rule_cat, rule_rate in rules.items():
        if rule_cat == "default":
            continue
        # Check if user's category or description contains the rule keyword
        # or rule keyword contains user's category (e.g. "Recharge" → "Recharge / Mobile")
        if (rule_cat.lower() in category_lower or rule_cat.lower() in desc_lower or
            category_lower in rule_cat.lower()):
            rate = rule_rate
            matched_rule = rule_cat
            break

    # Only calculate cashback for debit/expense transactions
    if amount <= 0:
        return 0.0, f"{matched_rule} (credit)"

    cashback = round(amount * rate, 2)
    label = f"{matched_rule} @ {rate*100:.0f}%"
    return cashback, label


def batch_calculate(transactions):
    """
    Calculate cashback for a list of transaction dicts.
    Adds 'cashback' and 'cashback_label' keys.
    """
    result = []
    total_cb = 0.0

    for txn in transactions:
        if "cashback" not in txn or txn.get("cashback") is None:
            cb, label = calculate_cashback(
                txn.get("amount", 0),
                txn.get("category", "Other"),
                txn.get("card_id", "other"),
                txn.get("description", ""),
            )
            txn["cashback"] = cb
            txn["cashback_label"] = label
        total_cb += txn.get("cashback", 0)
        result.append(txn)

    return result, round(total_cb, 2)


# Short human-readable label for each card
CARD_CB_LABELS = {
    "sbi_cb": "5% online (SBI CB)",
    "sbi_pp": "2% flat (PhonePe)",
    "hdfc_mil": "5% cat (Millennia)",
    "hdfc_swig": "10% food (Swiggy)",
    "bob_eterna": "2.5% (Eterna)",
    "other": "—",
    "other_card": "—",
}
