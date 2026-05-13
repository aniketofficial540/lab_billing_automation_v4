from num2words import num2words


def to_indian_words(amount: float) -> str:
    """Convert a numeric amount to Indian-format words.

    Examples:
        23190   → "Rupees Twenty Three Thousand One Hundred And Ninety Only"
        23190.5 → "Rupees Twenty Three Thousand One Hundred And Ninety And Fifty Paise Only"
    """
    try:
        amount = float(amount)
    except (TypeError, ValueError):
        raise ValueError(f"Cannot convert '{amount}' to words — not a valid number.")

    # Separate rupees and paise
    rupees = int(amount)
    paise_raw = round((amount - rupees) * 100)
    paise = int(paise_raw)

    rupee_words = num2words(rupees, lang="en_IN").strip()
    rupee_words = _title_case(rupee_words)

    if paise > 0:
        paise_words = num2words(paise, lang="en_IN").strip()
        paise_words = _title_case(paise_words)
        return f"Rupees {rupee_words} And {paise_words} Paise Only"

    return f"Rupees {rupee_words} Only"


def _title_case(text: str) -> str:
    """Remove commas/hyphens, title-case each word."""
    text = text.replace("-", " ").replace(",", "")
    return " ".join(w.capitalize() for w in text.split())
