/*
   Intrusive single/double linked dllist for C++.

   version 0.1, augusti, 2013

   Copyright (C) 2013- Fredrik Kihlander

   This software is provided 'as-is', without any express or implied
   warranty.  In no event will the authors be held liable for any damages
   arising from the use of this software.

   Permission is granted to anyone to use this software for any purpose,
   including commercial applications, and to alter it and redistribute it
   freely, subject to the following restrictions:

   1. The origin of this software must not be misrepresented; you must not
      claim that you wrote the original software. If you use this software
      in a product, an acknowledgment in the product documentation would be
      appreciated but is not required.
   2. Altered source versions must be plainly marked as such, and must not be
      misrepresented as being the original software.
   3. This notice may not be removed or altered from any source distribution.

   Fredrik Kihlander

   mboxview team: renamed file name and class names to avoid name clash
   Updated remove_head and remove-tail to return NULL if list empty.
   Added variable to keep count of elements on the list.
   TODO: introduce sentinel to simplify implementation ?
*/

#ifndef INTRUSIVE_LIST_H_INCLUDED
#define INTRUSIVE_LIST_H_INCLUDED

#include "assert.h"

#if defined( _INTRUSIVE_LIST_ASSERT_ENABLED )
#  if !defined( _INTRUSIVE_LIST_ASSERT )
#    include <assert.h>
#    define _INTRUSIVE_LIST_ASSERT( cond, ... ) assert( cond )
#  endif
#else
#  if !defined( _INTRUSIVE_LIST_ASSERT )
#    define _INTRUSIVE_LIST_ASSERT( ... )
#  endif
#endif

template <typename T>
struct dlink_node
{
	dlink_node()
		: next(0x0)
		, prev(0x0)
	{}

	T* next;
	T* prev;
};

/**
 * intrusive double linked dllist.
 */
template <typename T, dlink_node<T> T::*NODE>
class dllist
{
	T* head_ptr;
	T* tail_ptr;
	int cnt;

public:
	 dllist()
	 	 : head_ptr( 0x0 )
	 	 , tail_ptr( 0x0 )
		 , cnt(0)
	 {}
	~dllist() { clear(); }

	bool find(T* elem)
	{
		T *e;
		for (e = head_ptr; e != 0; (&(item->*NODE))->next)
		{
			if (e == elem)
				return true;
		}
		return false;
	}

	/**
	 * insert item at the head of dllist.
	 * @param item item to insert in dllist.
	 */
	void insert_head( T* elem )
	{
		dlink_node<T>* node = &(elem->*NODE);

		_INTRUSIVE_LIST_ASSERT( node->next == 0x0 );
		_INTRUSIVE_LIST_ASSERT( node->prev == 0x0 );

		node->prev = 0;
		node->next = head_ptr;

		if( head_ptr != 0x0 )
		{
			dlink_node<T>* last_head = &(head_ptr->*NODE);
			last_head->prev = elem;
		}

		head_ptr = elem;

		if( tail_ptr == 0x0 )
			tail_ptr = head_ptr;

		cnt++;
	}

	/**
	 * insert item at the tail of dllist.
	 * @param item item to insert in dllist.
	 */
	void insert_tail( T* item )
	{
		if( tail_ptr == 0x0 )
			insert_head( item );
		else
		{
			dlink_node<T>* tail_node = &(tail_ptr->*NODE);
			dlink_node<T>* item_node = &(item->*NODE);
			_INTRUSIVE_LIST_ASSERT( item_node->next == 0x0 );
			_INTRUSIVE_LIST_ASSERT( item_node->prev == 0x0 );
			tail_node->next = item;
			item_node->prev = tail_ptr;
			item_node->next = 0x0;
			tail_ptr = item;
			cnt++;
		}
	}

	/**
	 * insert item in dllist after other dllist item.
	 * @param item item to insert in dllist.
	 * @param in_list item that already is inserted in dllist.
	 */
	void insert_after( T* item, T* in_list )
	{
		dlink_node<T>* node    = &(item->*NODE);
		dlink_node<T>* in_node = &(in_list->*NODE);

		if( in_node->next )
		{
			dlink_node<T>* in_next = &(in_node->next->*NODE);
			in_next->prev = item;
		}

		node->next = in_node->next;
		node->prev = in_list;
		in_node->next = item;

		if( in_list == tail_ptr )
			tail_ptr = item;

		cnt++;
	}

	/**
	 * insert item in dllist before other dllist item.
	 * @param item item to insert in dllist.
	 * @param in_list item that already is inserted in dllist.
	 */
	void insert_before( T* item, T* in_list )
	{
		dlink_node<T>* node    = &(item->*NODE);
		dlink_node<T>* in_node = &(in_list->*NODE);

		if( in_node->prev )
		{
			dlink_node<T>* in_prev = &(in_node->prev->*NODE);
			in_prev->next = item;
		}

		node->next = in_list;
		node->prev = in_node->prev;
		in_node->prev = item;

		if( in_list == head_ptr )
			head_ptr = item;

		cnt++;
	}

	/**
	 * remove item from dllist.
	 * @param item the element to remove
	 * @note item must exist in dllist!
	 */
	void remove( T* item )
	{
		dlink_node<T>* node = &(item->*NODE);

		T* next = node->next;
		T* prev = node->prev;

		dlink_node<T>* next_node = &(next->*NODE);
		dlink_node<T>* prev_node = &(prev->*NODE);

		if( item == head_ptr ) head_ptr = next;
		if( item == tail_ptr ) tail_ptr = prev;
		if( prev != 0x0 ) prev_node->next = next;
		if( next != 0x0 ) next_node->prev = prev;

		node->next = 0x0;
		node->prev = 0x0;

		cnt--;
	}

	/**
	* remove the first item in the dllist.
	* @return the removed item.
	* mboxview team: NULL is returned if list empty
	*/
	T* remove_head()
	{
		T* ret = head();
		if (ret)
			remove(head());
		return ret;
	}

	/**
	* remove the last item in the dllist.
	* @return the removed item.
	* mboxview team: NULL is returned if list empty
	*/
	T* remove_tail()
	{
		T* ret = tail();
		if (ret)
			remove(tail());
		return ret;
	}

	/**
	 * return first item in dllist.
	 * @return first item in dllist or 0x0 if dllist is empty.
	 */
	T* head() { return head_ptr; }
	T* first() { return head_ptr; }

	/**
	 * return last item in dllist.
	 * @return last item in dllist or 0x0 if dllist is empty.
	 */
	T* tail() const { return tail_ptr; }
	T* last() const { return tail_ptr; }

	/**
	 * return next element in dllist after argument element or 0x0 on end of dllist.
	 * @param item item to get next element in dllist for.
	 * @note item must exist in dllist!
	 * @return element after item in dllist.
	 */
	T* next( T* item )
	{
		dlink_node<T>* node = &(item->*NODE);
		return node->next;
	}

	/**
	 * return next element in dllist after argument element or 0x0 on end of dllist.
	 * @param item item to get next element in dllist for.
	 * @note item must exist in dllist!
	 * @return element after item in dllist.
	 */
	const T* next( const T* item )
	{
		const dlink_node<T>* node = &(item->*NODE);
		return node->next;
	}

	/**
	 * return previous element in dllist after argument element or 0x0 on start of dllist.
	 * @param item item to get previous element in dllist for.
	 * @note item must exist in dllist!
	 * @return element before item in dllist.
	 */
	T* prev( T* item )
	{
		dlink_node<T>* node = &(item->*NODE);
		return node->prev;
	}

	/**
	 * return previous element in dllist after argument element or 0x0 on start of dllist.
	 * @param item item to get previous element in dllist for.
	 * @note item must exist in dllist!
	 * @return element before item in dllist.
	 */
	const T* prev( const T* item )
	{
		const dlink_node<T>* node = &(item->*NODE);
		return node->prev;
	}

	/**
	 * clear queue.
	 */
	void clear()
	{
		while( !empty() )
			remove( head() );

		assert(cnt == 0);
	}

	/**
	 * check if the dllist is empty.
	 * @return true if dllist is empty.
	 */
	bool empty() { return head_ptr == 0x0; }

	/**

 * @return number of elements on the list
 */
	int count() { return cnt; }
};

/**
 * macro to define node in struct/class to use in conjunction with dllist.
 *
 * @example
 * struct my_struct
 * {
 *     DLLIST_NODE( my_struct ) list1;
 *     DLLIST_NODE( my_struct ) list2;
 *
 *     int some_data;
 * };
 */
#define DLLIST_NODE( owner )    dlink_node<owner>

/**
 * macro to define dllist that act upon specific member in struct-type.
 *
 * @example
 * DLLIST( my_struct, list1 ) first_list;
 * DLLIST( my_struct, list2 ) second_list;
 */
#define DLLIST( owner, member ) dllist<owner, &owner::member>

#endif // INTRUSIVE_LIST_H_INCLUDED
