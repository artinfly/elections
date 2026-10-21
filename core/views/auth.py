"""Вход и выход из системы."""

from django.contrib.auth import authenticate, login, logout
from django.http import HttpRequest, HttpResponse
from django.shortcuts import redirect, render


def login_view(request: HttpRequest) -> HttpResponse:
    """Страница входа в систему (форма логина и пароля)."""
    if request.user.is_authenticated:
        return redirect("method")

    error = False
    if request.method == "POST":
        user = authenticate(
            request,
            username=request.POST.get("username"),
            password=request.POST.get("password"),
        )
        if user:
            login(request, user)
            return redirect("method")
        error = True

    return render(request, "login.html", {"error": error})


def logout_view(request: HttpRequest) -> HttpResponse:
    """
    Выход из системы — только POST-запросом (кнопка «Выйти» в меню).

    GET (адрес /logout/ открыли напрямую или по чужой ссылке) не разлогинивает,
    а возвращает на главную.
    """
    if request.method != "POST":
        return redirect("method")
    logout(request)
    return redirect("login")
