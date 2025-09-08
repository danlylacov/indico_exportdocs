#!/usr/bin/env bash
set -euo pipefail

# install_exportdocs_plugin.sh
# Установка плагина exportdocs для Indico из git-репозитория.
# Репозиторий: https://github.com/danlylacov/indico_exportdocs.git

# Параметры по умолчанию (меняйте под вашу инсталляцию):
INDICO_VENV_DEFAULT="/indico/env"
INDICO_CONF_DEFAULT="/dev/indico/src/indico/indico.conf"
PLUGIN_ENTRYPOINT_NAME="exportdocs"
GIT_URL_DEFAULT="https://github.com/danlylacov/indico_exportdocs.git"

usage() {
  cat <<EOF
Usage: $0 [--venv PATH] [--conf PATH] [--git URL] [--workdir DIR] [--restart]

Options:
  --venv PATH     Путь к virtualenv Indico (default: ${INDICO_VENV_DEFAULT})
  --conf PATH     Путь к indico.conf (default: ${INDICO_CONF_DEFAULT})
  --git URL       URL репозитория плагина (default: ${GIT_URL_DEFAULT})
  --workdir DIR   Каталог, где будет расположен код плагина (default: текущий каталог)
  --restart       Попытаться перезапустить Indico (systemd) после установки

Пример:
  sudo $0 --venv /opt/indico/.venv --conf /opt/indico/etc/indico.conf --restart
EOF
}

INDICO_VENV="${INDICO_VENV_DEFAULT}"
INDICO_CONF="${INDICO_CONF_DEFAULT}"
GIT_URL="${GIT_URL_DEFAULT}"
WORKDIR=""
DO_RESTART="0"

while [[ $# -gt 0 ]]; do
  case "$1" in
    --venv) INDICO_VENV="$2"; shift 2 ;;
    --conf) INDICO_CONF="$2"; shift 2 ;;
    --git) GIT_URL="$2"; shift 2 ;;
    --workdir) WORKDIR="$2"; shift 2 ;;
    --restart) DO_RESTART="1"; shift ;;
    -h|--help) usage; exit 0 ;;
    *) echo "Неизвестный параметр: $1"; usage; exit 1 ;;
  esac
done

if [[ -z "${WORKDIR}" ]]; then
  WORKDIR="$(pwd)"
fi

echo "[1/6] Проверка окружения..."
if [[ ! -x "${INDICO_VENV}/bin/python" || ! -x "${INDICO_VENV}/bin/pip" ]]; then
  echo "Ошибка: не найдено виртуальное окружение Indico по пути: ${INDICO_VENV}"
  exit 1
fi

if [[ ! -f "${INDICO_CONF}" ]]; then
  echo "Ошибка: не найден indico.conf по пути: ${INDICO_CONF}"
  exit 1
fi

echo "[2/6] Клонирование/обновление репозитория плагина..."
PLUGIN_DIR="${WORKDIR}/indico_exportdocs"
if [[ -d "${PLUGIN_DIR}/.git" ]]; then
  echo "Репозиторий уже существует: ${PLUGIN_DIR} — выполняю git pull"
  git -C "${PLUGIN_DIR}" fetch --all --prune
  git -C "${PLUGIN_DIR}" reset --hard origin/main
else
  git clone "${GIT_URL}" "${PLUGIN_DIR}"
fi

echo "[3/6] Генерация/обновление setup.py..."
SETUP_PY="${WORKDIR}/setup.py"
cat > "${SETUP_PY}" <<'PYSETUP'
from setuptools import setup

setup(
    name='indico-plugin-exportdocs',
    version='0.1.0',
    description='Indico plugin for exporting lists and reports as docx',
    author='ExportDocs',
    author_email='noreply@example.com',
    url='https://github.com/danlylacov/indico_exportdocs',
    packages=['indico_exportdocs'],
    install_requires=[
        'indico',
        'python-docx',
        'docxtpl',
    ],
    entry_points={
        'indico.plugins': [
            'exportdocs = indico_exportdocs.plugin:ExportDocsPlugin',
        ],
    },
    include_package_data=True,
    zip_safe=False,
    license='MIT',
)
PYSETUP

echo "[4/6] Установка плагина в venv (editable)..."
"${INDICO_VENV}/bin/pip" install -U pip wheel setuptools
"${INDICO_VENV}/bin/pip" install -e "${WORKDIR}"

echo "[5/6] Обновление indico.conf (добавление плагина в PLUGINS)..."

if grep -qE '^[[:space:]]*PLUGINS[[:space:]]*=' "${INDICO_CONF}"; then
  if grep -qE "PLUGINS[[:space:]]*=.*(['\"])${PLUGIN_ENTRYPOINT_NAME}\\1" "${INDICO_CONF}"; then
    echo "Плагин уже присутствует в PLUGINS."
  else
    TMP_CONF="$(mktemp)"
    python - "$INDICO_CONF" "$PLUGIN_ENTRYPOINT_NAME" > "$TMP_CONF" <<'PYEDIT'
import re, sys, json
path = sys.argv[1]
name = sys.argv[2]
data = open(path, 'r', encoding='utf-8').read()


def add_plugin(line):
    # Находим список между [ ... ]
    m = re.search(r'(PLUGINS\s*=\s*\[)(.*?)(\])', line, flags=re.S)
    if not m:
        return line
    prefix, content, suffix = m.groups()
    # Быстрое извлечение элементов (не идеально, но достаточно для обычных случаев)
    items = re.findall(r"(['\"])(.+?)\\1", content)
    values = [v for _, v in items]
    if name in values:
        return line
    values.append(name)
    # Сохраняем кавычки одинарные
    rendered = ', '.join("'" + v + "'" for v in values)
    return f"{prefix}{rendered}{suffix}"

out = []
in_plugins = False
for raw in data.splitlines(keepends=False):
    if re.match(r'^\s*PLUGINS\s*=', raw):
        in_plugins = True
        # возможно список многострочный — собираем пока не увидим ]
        buf = [raw]
        if ']' not in raw:
            # копим последующие строки
            continue
        else:
            # однострочный
            out.append(add_plugin(raw))
            in_plugins = False
    elif in_plugins:
        buf.append(raw)
        if ']' in raw:
            block = '\n'.join(buf)
            out.append(add_plugin(block))
            in_plugins = False
    else:
        out.append(raw)

# Если не нашли PLUGINS вовсе
if not any(re.match(r'^\s*PLUGINS\s*=', l) for l in out):
    out.append(f"PLUGINS = ['{name}']")

sys.stdout.write('\n'.join(out) + '\n')
PYEDIT
    cp "$TMP_CONF" "${INDICO_CONF}"
    rm -f "$TMP_CONF"
    echo "PLUGINS обновлён."
  fi
else
  echo "Секция PLUGINS не найдена — добавляю."
  printf "\nPLUGINS = ['%s']\n" "${PLUGIN_ENTRYPOINT_NAME}" >> "${INDICO_CONF}"
fi

echo "[6/6] Проверка установки..."
"${INDICO_VENV}/bin/python" - <<'PYCHK'
import pkg_resources
eps = list(pkg_resources.iter_entry_points('indico.plugins'))
names = [e.name for e in eps]
print("Entry points indico.plugins:", names)
assert 'exportdocs' in names, "Entry point 'exportdocs' не обнаружен"
print("OK: entry point 'exportdocs' найден")
PYCHK

if [[ "${DO_RESTART}" == "1" ]]; then
  echo "[Опционально] Перезапуск Indico через systemd..."
  if systemctl list-unit-files | grep -q '^indico\\.service'; then
    sudo systemctl restart indico.service
    echo "Indico перезапущен."
  else
    echo "Внимание: systemd unit indico.service не найден. Перезапустите ваш Indico вручную."
  fi
fi

echo "Готово. Плагин установлен."
