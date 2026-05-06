"""Utilitaires de cache pour le serveur MCP"""

import functools
from typing import Any, Callable


def simple_cache(func: Callable) -> Callable:
    """Décorateur de cache simple basé sur les arguments de la fonction.

    Utilise functools.lru_cache avec une taille illimitée.
    Compatible avec les méthodes d'instance (ignore `self` dans la clé de cache).
    """
    cache: dict = {}

    @functools.wraps(func)
    def wrapper(self, *args, **kwargs) -> Any:
        key = (args, tuple(sorted(kwargs.items())))
        if key not in cache:
            cache[key] = func(self, *args, **kwargs)
        return cache[key]

    wrapper.cache_clear = cache.clear  # type: ignore[attr-defined]
    return wrapper
