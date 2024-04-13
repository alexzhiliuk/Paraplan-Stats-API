# Paraplan Stats API

Python-скрипт для сбора статистики из системы Paraplan

## Установка

```bash
git clone <repository-link>
```

Используйте пакетный менеджер [pip](https://pip.pypa.io/en/stable/) для установки зависимостей.

```bash
pip install -r requirements.txt
```

## Переменные окружения

**LOGIN** - Логин от личного кабинет \
**PASS** - Пароль от личного кабинета

## Использование

Для генерации отчета “Непродленные абонементы за месяц”

```bash
python main.py current-month
```

Для генерации отчета “Непродленные абонементы за неделю”

```bash
python main.py current-week
```

Для генерации отчета “Прогноз учеников”

```bash
python main.py next-month
```