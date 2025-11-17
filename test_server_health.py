"""
Тестовый скрипт для проверки работы MCP сервера
Проверяет доступность сервера и отсутствие зависаний
"""

import requests
import time
import sys
from datetime import datetime

def test_server_health(host="127.0.0.1", port=8000, path="/mcp", num_requests=10, delay=0.5):
    """
    Тестирует работу сервера, отправляя несколько запросов
    
    Args:
        host: Хост сервера
        port: Порт сервера
        path: Путь к MCP endpoint
        num_requests: Количество запросов для отправки
        delay: Задержка между запросами в секундах
    """
    url = f"http://{host}:{port}{path}"
    print(f"Testing server at: {url}")
    print(f"Sending {num_requests} requests with {delay}s delay...\n")
    
    success_count = 0
    error_count = 0
    conflict_count = 0
    timeout_count = 0
    
    for i in range(1, num_requests + 1):
        try:
            timestamp = datetime.now().strftime("%H:%M:%S.%f")[:-3]
            print(f"[{timestamp}] Request {i}/{num_requests}...", end=" ")
            
            # Отправляем GET запрос с таймаутом
            response = requests.get(url, timeout=5)
            
            if response.status_code == 200:
                print(f"✓ OK (200)")
                success_count += 1
            elif response.status_code == 409:
                print(f"⚠ Conflict (409) - Server busy")
                conflict_count += 1
            else:
                print(f"✗ Error ({response.status_code})")
                error_count += 1
                
        except requests.exceptions.Timeout:
            print(f"✗ Timeout")
            timeout_count += 1
        except requests.exceptions.ConnectionError as e:
            print(f"✗ Connection Error: {e}")
            error_count += 1
        except Exception as e:
            print(f"✗ Unexpected error: {e}")
            error_count += 1
        
        # Задержка между запросами
        if i < num_requests:
            time.sleep(delay)
    
    # Итоги
    print("\n" + "="*50)
    print("RESULTS:")
    print(f"  ✓ Success:  {success_count}/{num_requests}")
    print(f"  ⚠ Conflicts: {conflict_count}/{num_requests}")
    print(f"  ✗ Errors:    {error_count}/{num_requests}")
    print(f"  ⏱ Timeouts:  {timeout_count}/{num_requests}")
    print("="*50)
    
    # Оценка здоровья сервера
    health_score = (success_count / num_requests) * 100
    print(f"\nHealth Score: {health_score:.1f}%")
    
    if health_score >= 90:
        print("Status: ✓ HEALTHY")
        return 0
    elif health_score >= 70:
        print("Status: ⚠ WARNING - Some issues detected")
        return 1
    else:
        print("Status: ✗ CRITICAL - Server has serious issues")
        return 2


def test_concurrent_requests(host="127.0.0.1", port=8000, path="/mcp", num_concurrent=5):
    """
    Тестирует обработку одновременных запросов
    """
    import concurrent.futures
    
    url = f"http://{host}:{port}{path}"
    print(f"\nTesting concurrent requests to: {url}")
    print(f"Sending {num_concurrent} concurrent requests...\n")
    
    def make_request(request_id):
        try:
            timestamp = datetime.now().strftime("%H:%M:%S.%f")[:-3]
            response = requests.get(url, timeout=10)
            status = "✓" if response.status_code == 200 else "✗"
            print(f"[{timestamp}] Request {request_id}: {status} ({response.status_code})")
            return response.status_code
        except Exception as e:
            print(f"Request {request_id}: ✗ Error - {e}")
            return None
    
    # Запускаем запросы параллельно
    with concurrent.futures.ThreadPoolExecutor(max_workers=num_concurrent) as executor:
        futures = [executor.submit(make_request, i) for i in range(1, num_concurrent + 1)]
        results = [f.result() for f in concurrent.futures.as_completed(futures)]
    
    success_count = sum(1 for r in results if r == 200)
    conflict_count = sum(1 for r in results if r == 409)
    
    print(f"\nConcurrent test results:")
    print(f"  ✓ Success:   {success_count}/{num_concurrent}")
    print(f"  ⚠ Conflicts: {conflict_count}/{num_concurrent}")
    
    if conflict_count > 0:
        print("\n⚠ Warning: Server is rejecting concurrent requests (409 Conflict)")
        print("   This may indicate that the server doesn't handle concurrent connections well.")


if __name__ == "__main__":
    # Параметры можно передать через командную строку
    host = sys.argv[1] if len(sys.argv) > 1 else "127.0.0.1"
    port = int(sys.argv[2]) if len(sys.argv) > 2 else 8000
    
    print("="*50)
    print("MCP Server Health Check")
    print("="*50)
    
    # Базовый тест
    exit_code = test_server_health(host=host, port=port, num_requests=10, delay=0.5)
    
    # Тест одновременных запросов
    test_concurrent_requests(host=host, port=port, num_concurrent=3)
    
    print("\nTest completed.")
    sys.exit(exit_code)

