//
//////////////////////////////////////////////////////////////////
//
//  Windows Mbox Viewer is a free tool to view, search and print mbox mail archives..
//
// Source code and executable can be downloaded from
//  https://sourceforge.net/projects/mbox-viewer/  and
//  https://github.com/eneam/mboxviewer
//
//  Copyright(C) 2019  Enea Mansutti, Zbigniew Minciel
//
//  This program is free software; you can redistribute it and/or modify
//  it under the terms of the version 3 of GNU Affero General Public License
//  as published by the Free Software Foundation; 
//
//  This program is distributed in the hope that it will be useful,
//  but WITHOUT ANY WARRANTY; without even the implied warranty of
//  MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.See the GNU
//  Library General Public License for more details.
//
//  You should have received a copy of the GNU Library General Public
//  License along with this program; if not, write to the
//  Free Software Foundation, Inc., 51 Franklin St, Fifth Floor,
//  Boston, MA  02110 - 1301, USA.
//
//////////////////////////////////////////////////////////////////
//

#pragma once

#include "dllist.h"

template <typename K, typename T, typename HASH, typename EQUAL, dlink_node<T> T::*NODE>
class IHashMap;

template <typename K, typename T, typename HASH, typename EQUAL, dlink_node<T> T::*NODE>
struct IHashMapIterator
{
	friend class IHashMap< K,  T,  HASH,  EQUAL, NODE>;

public:
	inline void clear()
	{
		element = nullptr;
		index = 0;
		list = nullptr;
	}

	IHashMapIterator() 
	{
		clear();
	}

	IHashMapIterator(IHashMapIterator &src)
	{
		if (this == &src)
			return;
		element = src.element;
		hashMap = src.hashMap;
		index = src.index;
		list = src.list;
	}

	T* element;
protected:
	IHashMap<K, T, HASH, EQUAL, NODE> *hashMap;
	int index;
	dllist<T, NODE> *list;
};

template <typename K, typename T, typename HASH, typename EQUAL, dlink_node<T> T::*NODE>
class IHashMap
{
public:
	using IHashMapIter = IHashMapIterator<K, T, HASH, EQUAL, NODE>;

	IHashMap(
		unsigned long size,
		HASH hash = HASH{},
		EQUAL equal = EQUAL{},
		bool makeSizeAsPrime = true)
		:
		m_hashMapSize(size),
		m_Hash(hash), 
		m_Equal(equal)
	{
		m_ar = new dllist<T, NODE>[m_hashMapSize];
		m_count = 0;
	};

	~IHashMap() { delete[] m_ar; };

protected:

	inline int get_map_count()
	{
		int count = 0;
		int list_count;
		dllist<T, NODE> *list = 0;
		int index;
		for (index = 0; index < m_hashMapSize; index++)
		{
			list = &m_ar[index];
			list_count = list->count();
			count += list_count;
		}
		return count;
	}

public:

	inline void insert(unsigned long hashsum, T *elem)
	{
		//int mcount;
		m_ar[hashsum%m_hashMapSize].insert_head(elem);
		m_count++;
		//assert((mcount = get_map_count()) == m_count);
	}

	inline void insert(K *key, T* elem)
	{
		insert(m_Hash(key), elem);
	}

	inline T *find(K *key, unsigned long hashsum) const
	{
		dllist<T, NODE> &list = m_ar[hashsum%m_hashMapSize];
		T *iter;
		for (iter = list.head(); iter != 0; iter = list.next(iter))
		{
			if (m_Equal(key, iter))
				return iter;
		}
		return nullptr;
	}

	inline T *find(K *key) const
	{
		return(find(key, m_Hash(key)));
	};

	inline T *remove(K *key, unsigned long hashsum)
	{
		dllist<T, NODE> &list = m_ar[hashsum%m_hashMapSize];
		T *iter;
		for (iter = list.head(); iter != 0; iter = list.next(iter))
		{
			if (m_Equal(key, iter))
			{
				list.remove(iter);
				m_count--;
				return iter;
			}
		}
		return nullptr;
	}

	inline T *remove(K *key)
	{
		return(remove(key, m_Hash(key)));
	};

	inline void clear()
	{
		int count_sv = m_count;
		dllist<T, NODE> *list = 0;
		int index;
		for (index = 0; index < m_hashMapSize; index++)
		{
			list = &m_ar[index];
			count_sv += list->count();
		}
		count_sv = m_count;
		for (index = 0; index < m_hashMapSize; index++)
		{
			list = &m_ar[index];
			count_sv -= list->count();
			list->clear();
		}
		assert(count_sv == 0);
		m_count = 0;
	}

	inline int count()
	{
		return m_count;
	}

protected:
	inline void find_next_non_empty_list(IHashMapIter &it)
	{
		assert(it.element == nullptr);
		dllist<T, NODE> *list = 0;
		it.element = nullptr;
		for (; it.index < m_hashMapSize; it.index++)
		{
			list = &m_ar[it.index];
			if ((it.element = list->head()) != nullptr)
			{
				it.list = list;
				return;
			}
		}
		assert(it.index == m_hashMapSize);
		it.list = &m_ar[it.index];
	}

	inline void next_helper(IHashMapIter &it)
	{
		dllist<T, NODE> *list = 0;

		if (it.element == nullptr)
		{
			find_next_non_empty_list(it);
		}
		else
		{
			assert(it.list != nullptr);
			it.element = list->next(it.element);
			if (it.element == nullptr)
			{
				it.index++;
				find_next_non_empty_list(it);
			}
		}
	}

public:

	// Iterator - not very similar to std iterator but it works
	// Allows to iterate and remove map elements
	// Only use void remove(IHashMapIter &it) to remove items while iterating
	//
	void remove(IHashMapIter &it)
	{
		assert(it.element != nullptr);
		assert(it.list != nullptr);
		T *elem = it.list->next(it.element);
		it.list->remove(it.element);
		m_count--;
		it.element = elem;
		if (it.element == nullptr)
		{
			it.index++;
			next_helper(it);
		}
	}

	inline IHashMapIter first()
	{
		IHashMapIter it;
		it.hashMap = this;
		next_helper(it);
		return it;
	}

	inline void next(IHashMapIter &it)
	{
		next_helper(it);
	}

	inline bool last(IHashMapIter &it)
	{
		dllist<T, NODE> *last_list = &m_ar[m_hashMapSize];

		if (it.list == last_list)
			return true;
		else
			return false;
	}

protected:
	dllist<T, NODE> *m_ar;
	unsigned long m_hashMapSize;
	int m_count;
	HASH m_Hash;
	EQUAL m_Equal;
};

#if 0

#include <string>

bool IHashMapTest()
{
	class Item
	{
	public:
		Item(std::string &nm) { name = nm; }
		//
		dlink_node<Item> hashNode;
		dlink_node<Item> hashNode2;
		std::string name;
	};

	struct ItemHash {
		unsigned long operator()(const Item *key) const
		{
			unsigned long  hashsum = std::hash<std::string>{}(key->name);
			return hashsum;
		};
		unsigned long operator()(Item *key) const
		{
			unsigned long  hashsum = std::hash<std::string>{}(key->name);
			return hashsum;
		};
	};

	struct ItemEqual {
		bool operator()(const Item *key, const Item *item) const
		{
			return (key->name == item->name);
		};
		bool operator()(Item *key, const Item *item) const
		{
			return (key->name == item->name);
		};
	};

	using MyHashMap = IHashMap<Item, Item, ItemHash, ItemEqual, &Item::hashNode>;
	using MyHashMap2 = IHashMap<Item, Item, ItemHash, ItemEqual, &Item::hashNode2>;

	unsigned long sz = 10;

	Item *it = new Item(std::string("John"));
	Item *it2 = new Item(std::string("Irena"));
	Item *it3 = new Item(std::string("Lucyna"));
	Item *it4 = new Item(std::string("Jan"));

	MyHashMap hashMap(sz);
	MyHashMap hashMap2(sz);

	/// Insert  it3 and it4 to second hash map hashMap2
	if (hashMap2.find(it3) == nullptr)
		hashMap2.insert(it3, it3);

	if (hashMap2.find(it4) == nullptr)
		hashMap2.insert(it4, it4);

	hashMap2.clear();
	assert(hashMap2.count() == 0);

	/// Insert it, it2 and it3
	if (hashMap.find(it) == nullptr)
		hashMap.insert(it, it);
	assert(hashMap.count() == 1);

	if (hashMap.find(it2) == nullptr)
		hashMap.insert(it2, it2);
	assert(hashMap.count() == 2);

	if (hashMap.find(it3) == nullptr)
		hashMap.insert(it3, it3);
	assert(hashMap.count() == 3);


	// Iterate and print item names
	MyHashMap::IHashMapIter iter = hashMap.first();
	for (; !hashMap.last(iter); hashMap.next(iter))
	{
		TRACE("%s\n", iter.element->name.c_str());
	}

	// Remove all items by Key one by one
	hashMap.remove(it);
	assert(hashMap.count() == 2);

	hashMap.remove(it2);
	assert(hashMap.count() == 1);

	hashMap.remove(it3);
	assert(hashMap.count() == 0);

	///Insert it, it 2 and i3 again to the same hash map
	if (hashMap.find(it) == nullptr)
		hashMap.insert(it, it);
	assert(hashMap.count() == 1);

	if (hashMap.find(it2) == nullptr)
		hashMap.insert(it2, it2);
	assert(hashMap.count() == 2);

	if (hashMap.find(it3) == nullptr)
		hashMap.insert(it3, it3);
	assert(hashMap.count() == 3);

	// Remove by key and hashsum
	Item *item;
	if ((item = hashMap.find(it2)) != nullptr)
	{
		unsigned long hashsum = ItemHash{}(it2);
		hashMap.remove(item, hashsum);
	}
	assert(hashMap.count() == 2);

	// Add back
	unsigned long hashsum = ItemHash{}(it2);
	if (hashMap.find(it2, hashsum) == nullptr)
		hashMap.insert(hashsum, it2);
	assert(hashMap.count() == 3);

	// Iterate and remove selected elements
	iter = hashMap.first();
	for (; !hashMap.last(iter); )
	{
		if (iter.element->name.compare("Lucyna") == 0)
		{
			TRACE("Keep %s\n", iter.element->name.c_str());
			hashMap.next(iter);
		}
		else
		{
			// Remove by Iterator
			TRACE("Remove %s\n", iter.element->name.c_str());
			hashMap.remove(iter);
		}
	}

	hashMap.clear();

	delete it;
	delete it2;
	delete it3;
	delete it4;

	int deb = 1;

	return true;
}

int next_prime(int L, int R)
{
	int i, j;
	if (L < 3) L = 3;

	for (i = L; i <= R; i++) 
	{
		for (j = 2; j <= i / 2; ++j) 
		{
			if (i % j == 0) 
			{
				break;
			}
		}
	}
	if (i < R)
		return i;
	else
		return L;
}
#endif





