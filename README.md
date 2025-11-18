```markdown
# Исследование спекания керамических материалов

**Автор:** Шатров Аскольд Игоревич

**Группа:** 4291

## Описание

Приложение исследует влияние давления газа и температуры спекания на плотность твёрдого сплава. Запуск на Ubuntu.

## Запуск на Ubuntu

### 1. Клонирование репозитория

```
git clone https://github.com/rightmelancholy/software_systems_development.git
cd software_systems_development
```

### 2. Создание виртуального окружения

```
python3 -m venv venv
source venv/bin/activate
```

### 3. Установка зависимостей

```
pip install numpy pandas matplotlib openpyxl
```

### 4. Запуск программы

```
python3 project.py
```

## Вход в систему

| Роль | Логин | Пароль |
|------|-------|--------|
| Исследователь | `researcher` | `pass123` |
| Администратор | `admin` | `admin123` |

## Файлы

- `project.py` — основной файл приложения
- `ceramics.db` — база данных (создаётся автоматически)
```

