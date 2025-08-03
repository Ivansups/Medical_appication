def calculate_prognosis(
    gender, age, weight, height, creatinine, creatinine_clearance,
    mpv, plcr, spontaneous_aggregation, induced_aggregation_1_ADP,
    induced_aggregation_5_ADP, induced_aggregation_15_ARA
):
    gender_num = 1 if gender == "Муж" else 2 if gender == "Жен" else 0

    const = -2.478
    k_gender = 0.477
    k_age = 0.05
    k_weight = -0.044
    k_height = -0.014
    k_creatinine = 0.05
    k_clearance = 0.063
    k_mpv = -0.448
    k_plcr = 0.029
    k_spont = 0.054
    k_ind1 = 0.012
    k_ind5 = -0.006
    k_ind15 = 0.027

    result = (
        const +
        k_gender * gender_num +
        k_age * (age or 0) +
        k_weight * (weight or 0) +
        k_height * (height or 0) +
        k_creatinine * (creatinine or 0) +
        k_clearance * (creatinine_clearance or 0) +
        k_mpv * (mpv or 0) +
        k_plcr * (plcr or 0) +
        k_spont * (spontaneous_aggregation or 0) +
        k_ind1 * (induced_aggregation_1_ADP or 0) +
        k_ind5 * (induced_aggregation_5_ADP or 0) +
        k_ind15 * (induced_aggregation_15_ARA or 0)
    )
    return result

def prognosis_text(result):
    if result < 1.56:
        return "Прогноз благоприятный, неблагоприятных событий в течение года не ожидается"
    elif 1.561 <= result <= 2.087:
        return "В течение года ожидается обращение к врачу-кардиологу"
    else:
        return "В течение ближайшего года вероятен летальный исход"