#!/bin/bash

set -e

RED='\033[0;31m'
GREEN='\033[0;32m'
YELLOW='\033[1;33m'
NC='\033[0m'

echo -e "${GREEN} Установка Reportsystem...${NC}"

if [ "$EUID" -eq 0 ]; then
  echo -e "${RED} Не запускайте скрипт от root! Запустите как обычный пользователь.${NC}"
  exit 1
fi

if ! command -v docker &> /dev/null; then
    echo -e "${YELLOW} Docker не найден. Устанавливаю...${NC}"
    curl -fsSL https://get.docker.com | sudo sh
    sudo usermod -aG docker $USER
    echo -e "${YELLOW}  Docker установлен. Применяю права без перезагрузки...${NC}"
    newgrp docker << ENDGRP
    REPO_DIR="reportsystem"
    if [ -d "$REPO_DIR" ]; then
        echo -e "${YELLOW}  Папка $REPO_DIR уже существует. Пропускаю клонирование.${NC}"
    else
        echo -e "${GREEN} Клонирование репозитория...${NC}"
        git clone https://github.com/ruslan-kulik/reportsystem.git $REPO_DIR
    fi

    cd $REPO_DIR

    echo -e "${GREEN} Сборка и запуск контейнеров (это может занять время)...${NC}"
    docker compose up -d --build

    echo -e "${GREEN}🗄 Применение миграций базы данных...${NC}"
    docker compose exec -T web python manage.py migrate --noinput

    echo -e "${GREEN} Сбор статических файлов...${NC}"
    docker compose exec -T web python manage.py collectstatic --noinput

    echo -e "${GREEN} Создание администратора...${NC}"
    echo -e "${YELLOW}Введите данные для суперпользователя (логин, email, пароль):${NC}"
    docker compose exec web python manage.py createsuperuser

    echo -e "${GREEN} Установка завершена!${NC}"
    echo -e "${GREEN} Откройте в браузере: http://localhost:8000${NC}"
    echo -e "${GREEN} Админка: http://localhost:8000/admin${NC}"
    echo ""
    echo -e "${YELLOW}Полезные команды:${NC}"
    echo "  • Остановить:   docker compose down"
    echo "  • Просмотр логов: docker compose logs -f web"
    echo "  • Обновить:     cd reportsystem && git pull && docker compose up -d --build"
    ENDGRP
else
    # Docker уже установлен — просто запускаем проект
    echo -e "${GREEN} Docker уже установлен.${NC}"

    REPO_DIR="reportsystem"
    if [ -d "$REPO_DIR" ]; then
        echo -e "${YELLOW}  Папка $REPO_DIR уже существует. Пропускаю клонирование.${NC}"
    else
        echo -e "${GREEN} Клонирование репозитория...${NC}"
        git clone https://github.com/ruslan-kulik/reportsystem.git $REPO_DIR
    fi

    cd $REPO_DIR

    echo -e "${GREEN} Сборка и запуск...${NC}"
    docker compose up -d --build

    echo -e "${GREEN} Применение миграций...${NC}"
    docker compose exec -T web python manage.py migrate --noinput

    echo -e "${GREEN} Создайте администратора вручную:${NC}"
    echo -e "${YELLOW}  docker compose exec web python manage.py createsuperuser${NC}"

    echo -e "${GREEN} Готово! Откройте: http://localhost:8000${NC}"
fi