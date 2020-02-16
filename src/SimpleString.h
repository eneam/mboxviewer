#pragma once

// Investigate and use standard CString instead  ?
class SimpleString
{
public:
#define DFLT_GROW_SIZE 16
	SimpleString() { Initialize(0, DFLT_GROW_SIZE); }
	SimpleString(int capacity) { Initialize(capacity, DFLT_GROW_SIZE); }
	SimpleString(int capacity, int grow_size) { Initialize(capacity, grow_size); }

	void Initialize(int capacity, int grow_size) {
		m_data = 0;  m_count = 0;  m_grow_size = grow_size;
		m_capacity = capacity;
		if (m_capacity < DFLT_GROW_SIZE)
			m_capacity = DFLT_GROW_SIZE;
		m_data = new char[m_capacity + 1];  // extra byte for NULL
		if (m_data)
			m_data[0] = 0;
		else
			m_capacity = 0;  // this would be a problem. introduce m_error ?
	};
	~SimpleString() { delete[] m_data; m_data = 0; m_count = 0; m_capacity = 0; };
	void Release() { delete[] m_data; m_data = 0; m_count = 0;  m_capacity = 0; };

	char *m_data;
	int m_capacity;
	int m_count;
	int m_grow_size;

	char *Data() { return m_data; }
	char *Data(int pos) { return &m_data[pos]; }  // zero based pos
	int Capacity() { return m_capacity; }
	int Count() { return m_count; }
	void SetCount(int count) {
		ASSERT(count <= m_capacity); ASSERT(m_data);
		m_count = count; m_data[count] = 0;
	}
	void Clear() {
		SetCount(0);
	}
	bool CanAdd(int characters) {
		if ((m_count + characters) <= m_capacity)
			return true; else  return false;
	}
	int Resize(int size);

	int ClearAndResize(int size) {
		int new_size = Resize(size);
		SetCount(0);
		return new_size;
	}

	void Copy(char c) {
		if (1 > m_capacity) // should never be true
			Resize(1);
		m_data[0] = c;
		SetCount(1);
	}

	void Append(char c) {
		if (!CanAdd(1)) Resize(m_capacity + 1);
		m_data[m_count++] = c;
		m_data[m_count] = 0;
	}

	void Copy(void const* Src, size_t  Size) {
		if (Size > (size_t)m_capacity) Resize(Size);  // Not good; you are mixing signed and unsigned
		::memcpy(m_data, Src, Size);
		SetCount(Size);
	}

	void append_internal(void const* Src, size_t  Size);

	void Append(void const* Src, size_t  Size) {
		if (!CanAdd(Size))
			Resize(m_count + Size);
		append_internal(Src, Size);
	}

	void Copy(register char *src) {
		int slen = strlen(src);
		SimpleString::Copy(src, slen);
	}

	void Append(register char *src) {
		int slen = strlen(src);
		SimpleString::Append(src, slen);
	}

	void Copy(SimpleString &str) {
		Copy(str.Data(), str.Count());
	}

	void Append(SimpleString &str) {
		Append(str.Data(), str.Count());
	}

	int FindNoCase(int offset, void const* Src, int  Size);
	int FindAny(int offset, void const * Src);
	int Find(int offset, char const c);

	char GetAt(int pos) {
		return m_data[pos];
	}
};

