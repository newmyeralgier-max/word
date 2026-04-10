import sys
import io

# Класс для фильтрации stdout от лишнего текста (баннеров и логов)
class SimpleFilteredStdout:
    def __init__(self, original):
        self.original = original
        self.encoding = getattr(original, 'encoding', 'utf-8')
        
    def write(self, s):
        # Отфильтровываем символы ASCII-арта и приветственные сообщения
        if any(c in s for c in ['╭', '│', '╰', '─']):
            return len(s)
        if "FastMCP server" in s:
            return len(s)
        return self.original.write(s)
        
    def flush(self):
        if hasattr(self.original, 'flush'):
            self.original.flush()
        
    def __getattr__(self, name):
        return getattr(self.original, name)

# Патчим stdout ДО импорта сервера, так как FastMCP инициализируется при загрузке модуля
original_stdout = sys.stdout
sys.stdout = SimpleFilteredStdout(original_stdout)

try:
    from notebooklm_mcp.server import main
except ImportError:
    print("Ошибка: Пакет notebooklm-mcp-server не установлен.", file=sys.stderr)
    sys.exit(1)

if __name__ == "__main__":
    sys.exit(main())
