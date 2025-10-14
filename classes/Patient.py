from dataclasses import dataclass, field
from enum import Enum
from typing import List, Optional


class Gender(str, Enum):
    MALE = "Муж"
    FEMALE = "Жен"


class CYP2C19(str, Enum):
    STAR1 = "CYP 2c19*1"
    STAR2 = "CYP 2c19*2"
    STAR3 = "CYP 2c19*3"
    STAR17 = "CYP 2c19*17"


class ABCB1(str, Enum):
    TT = "TT"
    TC = "TC"
    CC = "CC"


@dataclass
class PatientData:
    gender: Optional[Gender] = None
    age: Optional[int] = None
    T: Optional[float] = None
    weight: Optional[float] = None
    height: Optional[float] = None
    creatinine: Optional[float] = None
    creatinine_clearance: Optional[float] = None
    mpv: Optional[float] = None
    plcr: Optional[float] = None
    spontaneous_aggregation: Optional[float] = None
    induced_aggregation_1_ADP: Optional[float] = None
    induced_aggregation_5_ADP: Optional[float] = None
    induced_aggregation_15_ARA: Optional[float] = None
    cyp2c19: Optional[CYP2C19] = None
    abcb1: Optional[ABCB1] = None
    drugs: List[str] = field(default_factory=list)
    prognosis_coefficient: Optional[float] = None
    prognosis_result: Optional[str] = None
