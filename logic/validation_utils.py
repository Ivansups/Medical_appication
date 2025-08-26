def validate_age(age_text):
    """Валидация возраста"""
    try:
        age = int(age_text)
        if age <= 0 or age > 120:
            return False
        else:
            return True
    except ValueError:
        return False

def validate_weight(weight_text):
    """Валидация веса"""
    try:
        weight = float(weight_text)
        if weight <= 0 or weight > 300:
            return False
        else:
            return True
    except ValueError:
        return False

def validate_height(height_text):
    """Валидация роста"""
    try:
        height = float(height_text)
        if height <= 0 or height > 250:
            return False
        else:
            return True
    except ValueError:
        if height_text:
            return False
        else:
            return True

def validate_creatinine(creatinine_text):
    """Валидация креатинина"""
    try:
        creatinine = float(creatinine_text)
        if creatinine <= 0 or creatinine > 1000:
            return False
        else:
            return True
    except ValueError:
        if creatinine_text:
            return False
        else:
            return True

def validate_mpv(mpv_text):
    """Валидация MPV"""
    try:
        mpv = float(mpv_text)
        if mpv <= 0 or mpv > 20:
            return False
        else:
            return True
    except ValueError:
        if mpv_text:
            return False
        else:
            return True

def validate_plcr(plcr_text):
    """Валидация PLCR"""
    try:
        plcr = float(plcr_text)
        if plcr < 0 or plcr > 100:
            return False
        else:
            return True
    except ValueError:
        if plcr_text:
            return False
        else:
            return True

def validate_spontaneous_aggregation(agg_text):
    """Валидация спонтанной агрегации"""
    try:
        agg = float(agg_text)
        if agg < 0 or agg > 100:
            return False
        else:
            return True
    except ValueError:
        if agg_text:
            return False
        else:
            return True

def validate_induced_aggregation_1_ADP(agg_text):
    """Валидация индуцированной агрегации 1 мкМоль АДФ"""
    try:
        agg = float(agg_text)
        if agg < 0 or agg > 100:
            return False
        else:
            return True
    except ValueError:
        if agg_text:
            return False
        else:
            return True

def validate_induced_aggregation_5_ADP(agg_text):
    """Валидация индуцированной агрегации 5 мкМоль АДФ"""
    try:
        agg = float(agg_text)
        if agg < 0 or agg > 100:
            return False
        else:
            return True
    except ValueError:
        if agg_text:
            return False
        else:
            return True

def validate_induced_aggregation_15_ARA(agg_text):
    """Валидация индуцированной агрегации 15 мкл арахидоновой кислоты"""
    try:
        agg = float(agg_text)
        if agg < 0 or agg > 100:
            return False
        else:
            return True
    except ValueError:
        if agg_text:
            return False
        else:
            return True

def validate_platelet_count(platelets_text):
    """Валидация количества тромбоцитов"""
    try:
        platelets = float(platelets_text)
        if platelets <= 0 or platelets > 1000:
            return False
        else:
            return True
    except ValueError:
        if platelets_text:
            return False
        else:
            return True

def get_drug_cancellation_recommendation(platelet_count, drug_type):
    """Получить рекомендацию по отмене препарата на основе уровня тромбоцитов"""
    try:
        platelets = float(platelet_count)
        if platelets <= 10:
            if drug_type == "АСК":
                return "Рекомендовано отменить АСК"
            elif drug_type == "АСК+тикагрелор":
                return "Рекомендовано отменить тикагрелор и АСК"
            elif drug_type == "АСК+клопидогрел":
                return "Рекомендовано отменить клопидогрел и АСК"
            elif drug_type == "клопидогрел":
                return "Рекомендовано отменить клопидогрел"
        elif 10 < platelets <= 30:
            if drug_type in ["клопидогрел", "АСК+клопидогрел"]:
                return "Рекомендовано отменить клопидогрел"
            elif drug_type in ["АСК+тикагрелор", "Тикагрелор"]:
                return "Рекомендовано отменить тикагрелор"
            else:
                return "Прием может быть продолжен"
        elif 30 < platelets <= 50: 
            if drug_type in ["АСК+тикагрелор", "Тикагрелор"]:
                return "Рекомендовано отменить тикагрелор"
            elif drug_type in ["клопидогрел", "АСК", "АСК+клопидогрел"]:
                return "Прием может быть продолжен"
            else: 
                return "Прием может быть продолжен"
        else:
            return "Прием может быть продолжен"
    except:
        return "Не определено"
