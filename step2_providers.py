# -*- coding: utf-8 -*-
"""
Provider/Registry para Etapa 2 (Recebíveis).

Objetivo:
- Isolar lógicas específicas por bandeira sem quebrar as outras.
- Permitir adicionar novas bandeiras no futuro com mínimo impacto.

Regra:
- Por enquanto, cada provider apenas DELEGA para o runner existente do core,
  para manter comportamento atual.
"""

from dataclasses import dataclass
from typing import Callable, Optional, Any


RunnerFn = Callable[..., Any]


@dataclass
class BaseStep2Provider:
    code: str
    runner: RunnerFn  # função existente do core que executa Etapa 2

    def run_step2(self, conn, prov: str, month: str, d1: Optional[str], d2: Optional[str], window_days: int = 0):
        # Delegação: mantém o comportamento atual do core
        return self.runner(conn, prov, month, d1, d2, window_days=window_days)


class AleloProvider(BaseStep2Provider):
    pass


class TicketProvider(BaseStep2Provider):
    pass


class FarmaciasAppProvider(BaseStep2Provider):
    pass


def get_provider(prov_code: str, runner: RunnerFn) -> BaseStep2Provider:
    p = (prov_code or "").strip().upper()
    if p == "ALELO":
        return AleloProvider(code="ALELO", runner=runner)
    if p == "TICKET":
        return TicketProvider(code="TICKET", runner=runner)
    if p in ("FARMACIASAPP", "FARMACIAS_APP", "FARMACIAS APP"):
        return FarmaciasAppProvider(code="FARMACIASAPP", runner=runner)
    # fallback: mantém tudo funcionando mesmo se vier novo prov sem provider ainda
    return BaseStep2Provider(code=p, runner=runner)
