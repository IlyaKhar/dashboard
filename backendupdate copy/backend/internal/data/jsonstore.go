package data

import (
	"encoding/json"
	"fmt"
	"os"
	"sync"
	"time"
)

type cacheEntry struct {
	modTime time.Time
	raw     []byte
}

// JSONStore отвечает за чтение JSON-файлов с простым in-memory кешированием.
// Ключ кеша — абсолютный путь к файлу + время модификации.
type JSONStore struct {
	mu    sync.Mutex
	cache map[string]cacheEntry
}

func NewJSONStore() *JSONStore {
	return &JSONStore{
		cache: make(map[string]cacheEntry),
	}
}

// loadRaw читает файл с учётом кеша и возвращает байты JSON.
func (s *JSONStore) loadRaw(path string) ([]byte, error) {
	if path == "" {
		return nil, fmt.Errorf("jsonstore: пустой путь к файлу")
	}

	info, err := os.Stat(path)
	if err != nil {
		if os.IsNotExist(err) {
			return nil, fmt.Errorf("jsonstore: файл не найден: %s", path)
		}
		return nil, fmt.Errorf("jsonstore: ошибка доступа к файлу %s: %w", path, err)
	}

	modTime := info.ModTime()

	s.mu.Lock()
	entry, ok := s.cache[path]
	s.mu.Unlock()

	if ok && entry.modTime.Equal(modTime) && len(entry.raw) > 0 {
		return entry.raw, nil
	}

	raw, err := os.ReadFile(path)
	if err != nil {
		return nil, fmt.Errorf("jsonstore: ошибка чтения файла %s: %w", path, err)
	}

	s.mu.Lock()
	s.cache[path] = cacheEntry{
		modTime: modTime,
		raw:     raw,
	}
	s.mu.Unlock()

	return raw, nil
}

// LoadJSON читает JSON-файл по указанному пути и мапит его в целевой тип T.
// Если файл уже был прочитан и не менялся, используется кешированное содержимое.
func LoadJSON[T any](store *JSONStore, path string) (T, error) {
	var zero T

	if store == nil {
		return zero, fmt.Errorf("jsonstore: не инициализирован store")
	}

	raw, err := store.loadRaw(path)
	if err != nil {
		return zero, err
	}

	var out T
	if err := json.Unmarshal(raw, &out); err != nil {
		return zero, fmt.Errorf("jsonstore: ошибка парсинга JSON %s: %w", path, err)
	}

	return out, nil
}

