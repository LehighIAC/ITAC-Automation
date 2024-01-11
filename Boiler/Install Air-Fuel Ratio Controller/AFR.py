def AFR(CAT: int, FGT: int, XO2: float) -> float:
    """
    Extracted from Algorithm Document for PHASTEx
    Returns available heat, %
    param CAT: combution air temperature, degF
    param FGT: flue gas temperature, degF
    param XO2: excessive oxygen, %
    param XAir: excessive air, %
    """
    if XO2 < 0 or XO2 > 22:
        raise Exception("Excessive oxygen is invalid.")
    XO2 /= 100
    XAir = 8.52381 * XO2 / (2 - (9.52381 * XO2))
    Cp = 0.0178285179931519 + 0.00000255632 * FGT
    Heat = 95 - 0.025 * CAT
    XAirCorr = -(-1.078914 + Cp * CAT) * XAir
    PhtAirCorr = (-1.078914 + Cp * FGT) * (1 + XAir)
    AH = Heat + XAirCorr + PhtAirCorr
    if AH < 0 or AH > 100:
        raise Exception("Algorithm error!")
    else:
        AH = round(AH, 1)
        return AH