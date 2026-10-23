import secrets

MIN_PASSWORD_LENGTH = 8
MAX_PASSWORD_LENGTH = 256

DIGITS = "0123456789"
LOWERCASE = "abcdefghijkmnopqrstuvwxyz"
UPPERCASE = "ABCDEFGHIJKMNOPQRSTUVWXYZ"
SYMBOLS = "@#$%=:?./|~>*()<"


def generate_secure_password(length: int = 12) -> str:
    """Generate a random password from the operating system's secure random source.

    The password holds at least one digit, uppercase letter, lowercase letter
    and symbol; the remaining characters are drawn from all four sets.

    Args:
        length: The length of the password to generate, 8 to 256 (default: 12)

    Returns:
        A random password string

    Raises:
        ValueError: when length is outside 8 to 256
    """
    if not MIN_PASSWORD_LENGTH <= length <= MAX_PASSWORD_LENGTH:
        raise ValueError(
            f"The password length must be {MIN_PASSWORD_LENGTH} to {MAX_PASSWORD_LENGTH} characters, got {length}"
        )

    combined = DIGITS + UPPERCASE + LOWERCASE + SYMBOLS
    chars = [secrets.choice(charset) for charset in (DIGITS, UPPERCASE, LOWERCASE, SYMBOLS)]
    chars += [secrets.choice(combined) for _ in range(length - len(chars))]
    # Shuffle so the four guaranteed characters do not sit at the start.
    secrets.SystemRandom().shuffle(chars)
    return "".join(chars)
