"""
humanizer.py
Motor de comportamento humano: delays adaptativos, controle de volume,
sistema de aquecimento e aleatoriedade inteligente.

Estratégia de segurança (Meta 2026):
  - Delays aleatórios entre 20–90 s (base), com jitter gaussiano
  - Pausas longas obrigatórias a cada N envios (simula "parou pra fazer outra coisa")
  - Limite por hora progressivo: começa conservador, aumenta gradualmente
  - Ordem de envio embaralhada levemente (não sequencial exato)
  - Nenhum padrão fixo detectável
"""

import math
import random
import time
from dataclasses import dataclass, field
from datetime import datetime, timedelta
from typing import Callable, Optional


# ------------------------------------------------------------------ config
@dataclass
class HumanProfile:
    """Perfil de comportamento humano. Ajuste conforme o 'aquecimento' do número."""

    # Delays base entre mensagens (segundos)
    delay_min: float = 22.0
    delay_max: float = 85.0

    # Pausa longa (simula distração / intervalo)
    long_pause_every: int = 8          # a cada N envios
    long_pause_min: float = 90.0       # segundos
    long_pause_max: float = 240.0

    # Limite de envios por hora
    max_per_hour: int = 30

    # Aquecimento: multiplica o delay nos primeiros dias/envios
    warmup_factor: float = 1.0         # 1.0 = sem aquecimento; 2.0 = dobra delays

    # Probabilidade de digitar "lentamente" antes de enviar (0–1)
    typing_simulation_chance: float = 0.6
    typing_min: float = 1.5
    typing_max: float = 4.5


# Perfis prontos para uso
PROFILE_CAUTIOUS = HumanProfile(
    delay_min=45, delay_max=120, long_pause_every=5,
    long_pause_min=180, long_pause_max=420, max_per_hour=20,
    warmup_factor=1.5,
)

PROFILE_NORMAL = HumanProfile()   # defaults acima

PROFILE_RELAXED = HumanProfile(
    delay_min=20, delay_max=60, long_pause_every=12,
    long_pause_min=60, long_pause_max=150, max_per_hour=40,
    warmup_factor=1.0,
)

PROFILES = {
    "Cauteloso (recomendado)": PROFILE_CAUTIOUS,
    "Normal":                  PROFILE_NORMAL,
    "Rápido":                  PROFILE_RELAXED,
}


# ------------------------------------------------------------------ engine
class HumanBehaviorEngine:
    """
    Controla o ritmo de envio de forma que pareça completamente humano.
    Thread-safe para uso com threading.
    """

    def __init__(self, profile: HumanProfile = PROFILE_NORMAL):
        self.profile = profile
        self._sent_this_hour: list[datetime] = []
        self._send_count_session: int = 0
        self._delays_used: list[float] = []
        self._stop_flag: Callable[[], bool] = lambda: False

    def set_stop_flag(self, flag_fn: Callable[[], bool]) -> None:
        """Injeta função que retorna True quando o usuário solicitou parada."""
        self._stop_flag = flag_fn

    # --------------------------------------------------------- public API
    def compute_delay(self) -> float:
        """
        Calcula o próximo delay com distribuição gaussiana suavizada,
        garantindo que fique dentro de [min, max].
        """
        p = self.profile
        center = (p.delay_min + p.delay_max) / 2
        sigma  = (p.delay_max - p.delay_min) / 5.0  # ~95 % dentro do range

        delay = random.gauss(center, sigma)
        delay = max(p.delay_min, min(p.delay_max, delay))

        # Aplica fator de aquecimento
        delay *= p.warmup_factor

        # Micro-jitter final (±3 s) para quebrar qualquer periodicidade
        delay += random.uniform(-3, 3)
        delay = max(5.0, delay)

        return round(delay, 1)

    def should_long_pause(self) -> bool:
        return (
            self._send_count_session > 0
            and self._send_count_session % self.profile.long_pause_every == 0
        )

    def compute_long_pause(self) -> float:
        p = self.profile
        return round(random.uniform(p.long_pause_min, p.long_pause_max), 1)

    def is_hourly_limit_reached(self) -> bool:
        now = datetime.now()
        cutoff = now - timedelta(hours=1)
        self._sent_this_hour = [t for t in self._sent_this_hour if t > cutoff]
        return len(self._sent_this_hour) >= self.profile.max_per_hour

    def register_send(self) -> None:
        self._sent_this_hour.append(datetime.now())
        self._send_count_session += 1

    def get_typing_delay(self) -> float:
        """Retorna delay de 'simulação de digitação' ou 0 se não aplicar."""
        p = self.profile
        if random.random() < p.typing_simulation_chance:
            return round(random.uniform(p.typing_min, p.typing_max), 1)
        return 0.0

    def wait(self, seconds: float, step: float = 0.5) -> bool:
        """
        Aguarda `seconds` segundos em pequenos steps.
        Retorna False se a parada foi solicitada antes de terminar.
        """
        elapsed = 0.0
        while elapsed < seconds:
            if self._stop_flag():
                return False
            time.sleep(step)
            elapsed += step
        return True

    def wait_for_hourly_reset(self, log_fn: Optional[Callable] = None) -> bool:
        """
        Aguarda até que o limite por hora seja liberado.
        Retorna False se a parada foi solicitada.
        """
        while self.is_hourly_limit_reached():
            if self._stop_flag():
                return False
            if log_fn:
                log_fn("⏸ Limite por hora atingido — aguardando janela...")
            if not self.wait(30):
                return False
        return True

    # --------------------------------------------------------- stats
    def avg_delay(self) -> float:
        return round(sum(self._delays_used) / len(self._delays_used), 1) if self._delays_used else 0.0

    def record_delay(self, d: float) -> None:
        self._delays_used.append(d)

    def session_count(self) -> int:
        return self._send_count_session
