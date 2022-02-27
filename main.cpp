
#include <iostream>
#include<windows.h>
#include<conio.h>
#include<iomanip>
#include<fstream>
#include<string>
#include<vector>
using namespace std;

struct Cordinate
{
	int row = 0;
	int col = 0;
};

class Excel
{
	class Cell
	{
		friend class Excel;
		string data;
		Cell* up;
		Cell* down;
		Cell* left;
		Cell* right;
	public:
		Cell(string input = " ", Cell* l = nullptr, Cell* r = nullptr, Cell* u = nullptr, Cell* d = nullptr)
		{
			data = input;
			left = l;
			right = r;
			up = u;
			down = d;
		}
	};

	Cell* head;
	Cell* current;
	Cell* RangeStart;
	vector<vector<string>> Clipboard;
	Cordinate Range_start{};
	Cordinate Range_end{};
	int row_size, col_size;
	int c_row, c_col;

	int Cell_width = 10;
	int Cell_height = 4;

	void color(int k)
	{
		HANDLE hConsole = GetStdHandle(STD_OUTPUT_HANDLE);
		SetConsoleTextAttribute(hConsole, k);
		if (k > 255)
		{
			k = 1;
		}
	}

	void getRowColbyLeftClick(int& rpos, int& cpos)
	{
		HANDLE hInput = GetStdHandle(STD_INPUT_HANDLE);
		DWORD Events;
		INPUT_RECORD InputRecord;
		SetConsoleMode(hInput, ENABLE_PROCESSED_INPUT | ENABLE_MOUSE_INPUT | ENABLE_EXTENDED_FLAGS);
		do
		{
			ReadConsoleInput(hInput, &InputRecord, 1, &Events);
			if (InputRecord.Event.MouseEvent.dwButtonState == FROM_LEFT_1ST_BUTTON_PRESSED)
			{
				cpos = InputRecord.Event.MouseEvent.dwMousePosition.X;
				rpos = InputRecord.Event.MouseEvent.dwMousePosition.Y;
				break;
			}
		} while (true);
	}

	void gotoRowCol(int rpos, int cpos)
	{
		COORD scrn;
		HANDLE hOuput = GetStdHandle(STD_OUTPUT_HANDLE);
		scrn.X = cpos;
		scrn.Y = rpos;
		SetConsoleCursorPosition(hOuput, scrn);
	}

	void Print_cell(int row, int col, int colour)
	{
		color(colour);
		char c = '*';
		gotoRowCol(row * Cell_height, col * Cell_width);

		for (int i = 0; i < Cell_width; i++)
		{
			cout << c;
		}

		gotoRowCol(row * Cell_height + Cell_height, col * Cell_width);
		for (int i = 0; i <= Cell_width; i++)
		{
			cout << c;
		}


		for (int i = 0; i < Cell_height; i++)
		{

			gotoRowCol(row * Cell_height + i, col * Cell_width);
			cout << c;
		}

		int r = row * Cell_height;
		int ci = (col * Cell_width) + Cell_width;
		for (int i = 0; i <= Cell_height; i++)
		{

			gotoRowCol(r + i, ci);
			cout << c;
		}

		gotoRowCol((Cell_height * row) + Cell_height / 2, (col * Cell_width) + Cell_width / 2);
		cout << "     ";

	}

	void Print_cell_data(int row, int col, Cell* d, int colour)
	{
		color(colour);
		gotoRowCol((Cell_height * row) + Cell_height / 2, (col * Cell_width) + Cell_width / 2);
		cout << d->data;
	}

	void Print_Col()
	{
		for (int ri = 0, ci = c_col + 1; ri < row_size; ri++)
		{
			Print_cell(ri, ci, 7);
		}
	}

	void Print_Row()
	{
		for (int ci = 0, ri = c_row + 1; ci < col_size; ci++)
		{
			Print_cell(ri, ci, 7);
		}

	}

public:

	class iterator
	{
		Cell* t;
		friend class Excel;

		iterator()
		{
			t = nullptr;
		}

		iterator(Cell* data)
		{
			t = data;
		}

		iterator operator++()
		{
			if (t->right != nullptr)
				t = t->right;
			return *this;
		}

		iterator operator++(int)
		{
			if (t->down != nullptr)
				t = t->down;
			return *this;
		}

		iterator operator--()
		{
			if (t->left != nullptr)
				t = t->left;
			return *this;
		}

		iterator operator--(int)
		{
			if (t->right != nullptr)
				t = t->right;
			return *this;
		}

		bool operator==(iterator temp)
		{
			return (t == temp.t);
		}

		bool operator!=(iterator temp)
		{
			return (t != temp.t);
		}

		string& operator*()
		{
			return t->data;
		}

	};

	iterator Get_Head()
	{
		return iterator(head);
	}

	Excel()
	{
		head = nullptr;
		current = nullptr;
		row_size = col_size = 5;
		c_row = c_col = 0;

		head = New_row();
		current = head;
		for (int i = 0; i < row_size - 1; i++)
		{
			InsertRowBelow();
			row_size--;
			current = current->down;
		}


		current = head;

		Print_Grid();
		Print_Data();
	}

	Cell* New_row()
	{
		Cell* temp = new Cell();
		Cell* curr = temp;
		for (int i = 0; i < col_size - 1; i++)
		{
			Cell* temp2 = new Cell();
			temp->right = temp2;
			temp2->left = temp;
			temp = temp2;
		}

		return curr;
	}

	Cell* New_col()
	{
		Cell* temp = new Cell();
		Cell* curr = temp;
		for (int i = 0; i < row_size - 1; i++)
		{
			Cell* temp2 = new Cell();
			temp->down = temp2;
			temp2->up = temp;

			temp = temp2;
		}

		return curr;
	}

	Cell* InsertCellRight(Cell* curr, string data)
	{
		Cell* temp = new Cell(data);
		temp->left = curr;

		if (curr->right != nullptr)
		{
			temp->right = curr->right;
			temp->right->left = temp;
		}

		curr->right = temp;

		if (curr->up != nullptr && curr->up->right != nullptr)
		{
			temp->up = curr->up->right;
			curr->up->right->down = temp;
		}

		if (curr->down != nullptr && curr->down->right != nullptr)
		{
			temp->down = curr->down->right;
			curr->down->right->up = temp;
		}

		return temp;
	}

	Cell* InsertCellLeft(Cell* curr)
	{
		Cell* temp = new Cell("7");
		temp->right = curr;

		if (curr->left != nullptr)
		{
			temp->left = curr->left;
			temp->left->right = temp;
		}

		curr->left = temp;

		if (curr->up != nullptr && curr->up->left != nullptr)
		{
			temp->up = curr->up->right;
			curr->up->right->down = temp;
		}

		if (curr->down != nullptr && curr->down->left != nullptr)
		{
			temp->down = curr->down->left;
			curr->down->left->up = temp;
		}

		return temp;
	}

	void InsertColRight()
	{
		Cell* temp = New_col();
		Cell* temp2 = current;

		while (temp2->up != nullptr)
		{
			temp2 = temp2->up;
		}

		if (temp2->right == nullptr)
		{
			while (temp2 != nullptr)
			{
				temp2->right = temp;
				temp->left = temp2;

				temp2 = temp2->down;
				temp = temp->down;
			}
		}
		else
		{
			while (temp2 != nullptr)
			{
				temp->right = temp2->right;
				temp2->right = temp;
				temp->left = temp2;
				temp->right->left = temp;

				temp2 = temp2->down;
				temp = temp->down;
			}
		}

		col_size++;

	}

	void InsertColLeft()
	{

		Cell* temp = New_col();
		Cell* temp2 = current;


		while (temp2->up != nullptr)
		{
			temp2 = temp2->up;
		}

		if (temp2 == head)
		{
			head = temp;
		}
		if (temp2->left == nullptr)
		{
			while (temp2 != nullptr)
			{
				temp2->left = temp;
				temp->right = temp2;

				temp2 = temp2->down;
				temp = temp->down;
			}
		}
		else
		{
			while (temp2 != nullptr)
			{
				temp->left = temp2->left;
				temp2->left = temp;
				temp->right = temp2;
				temp->left->right = temp;

				temp2 = temp2->down;
				temp = temp->down;
			}
		}

		col_size++;
		Print_Grid();
		Print_Data();
	}

	void InsertRowBelow()
	{
		Cell* temp = New_row();
		Cell* temp2 = current;

		while (temp2->left != nullptr)
		{
			temp2 = temp2->left;
		}

		if (temp2->down == nullptr)
		{
			while (temp2 != nullptr)
			{
				temp2->down = temp;
				temp->up = temp2;

				temp = temp->right;
				temp2 = temp2->right;
			}
		}
		else
		{
			while (temp2 != nullptr)
			{
				temp->down = temp2->down;
				temp2->down = temp;
				temp->up = temp2;
				temp->down->up = temp;

				temp = temp->right;
				temp2 = temp2->right;
			}
		}

		row_size++;
	}

	void InsertRowAbove()
	{

		Cell* temp = New_row();
		Cell* temp2 = current;


		while (temp2->left != nullptr)
		{
			temp2 = temp2->left;
		}

		if (temp2 == head)
		{
			head = temp;
		}

		if (temp2->up == nullptr)
		{
			while (temp2 != nullptr)
			{
				temp2->up = temp;
				temp->down = temp2;

				temp = temp->right;
				temp2 = temp2->right;
			}
		}
		else
		{
			while (temp2 != nullptr)
			{
				temp->up = temp2->up;
				temp2->up = temp;
				temp->down = temp2;
				temp->up->down = temp;

				temp = temp->right;
				temp2 = temp2->right;
			}
		}
		row_size++;

		Print_Grid();
		Print_Data();
	}

	void InsertCellByRightShift()
	{
		Cell* temp = current;
		while (current->right != nullptr)
		{
			current = current->right;
		}
		InsertColRight();
		current = current->right;

		while (current != temp)
		{
			current->data = current->left->data;
			current = current->left;
		}
		current->data = "   ";
	}

	void InsertCellByLeftShift()
	{
		Cell* temp = current;
		while (current->left != nullptr)
		{
			current = current->left;
		}
		InsertColLeft();
		current = current->left;

		while (current != temp)
		{
			current->data = current->right->data;
			current = current->right;
		}
		current->data = " ";
		c_col++;

	}

	void InsertCellByDownShift()
	{
		Cell* temp = current;
		while (current->down != nullptr)
		{
			current = current->down;
		}
		InsertRowBelow();
		current = current->down;

		while (current != temp)
		{
			current->data = current->up->data;
			current = current->up;
		}
		current->data = " ";
	}

	void InsertCellByUpShift()
	{
		Cell* temp = current;
		while (current->up != nullptr)
		{
			current = current->up;
		}
		InsertRowAbove();
		current = current->up;

		while (current != temp)
		{
			current->data = current->down->data;
			current = current->down;
		}
		current->data = " ";
		c_row++;
	}

	void DeleteCellByLeftShift()
	{
		Cell* temp = current;
		temp->data = "    ";
		while (temp->right!= nullptr)
		{
			temp->data = temp->right->data;
			temp = temp->right;
		}
		temp->data = "    ";
	}

	void DeleteCellByUpShift()
	{
		Cell* temp = current;
		temp->data = " ";
		while (temp->down != nullptr)
		{
			temp->data = temp->down->data;
			temp = temp->down;
		}
		temp->data = " ";
	}

	void Delete_Col()
	{
		if (col_size <= 1)
			return;
		Cell* temp = current;

		while (temp->up != nullptr)
		{
			temp = temp->up;
		}

		Cell* delete_cell;

		if (temp == head)
		{
			head = temp->right;
		}


		if (temp->left == nullptr)
		{
			current = current->right;
			while (temp != nullptr)
			{
				delete_cell = temp;
				temp->right->left = nullptr;

				temp = temp->down;
				delete delete_cell;
			}

		}
		else if (temp->right == nullptr)
		{
			c_col--;
			current = current->left;
			while (temp != nullptr)
			{
				delete_cell = temp;
				temp->left->right = nullptr;
				temp = temp->down;
				delete delete_cell;
			}
		}
		else
		{
			c_col--;
			current = current->left;
			while (temp != nullptr)
			{
				delete_cell = temp;
				temp->left->right = temp->right;
				temp->right->left = temp->left;

				temp = temp->down;
				delete delete_cell;
			}
		}
		col_size--;
	}

	void Delete_Row()
	{

		if (row_size <= 1)
			return;
		Cell* temp = current;

		while (temp->left != nullptr)
		{
			temp = temp->left;
		}

		Cell* delete_cell;

		if (temp == head)
		{
			head = temp->down;
		}


		if (temp->up == nullptr)
		{
			current = current->down;
			while (temp != nullptr)
			{
				delete_cell = temp;
				temp->down->up = nullptr;

				temp = temp->right;
				delete delete_cell;
			}

		}
		else if (temp->down == nullptr)
		{
			c_row--;
			current = current->up;
			while (temp != nullptr)
			{
				delete_cell = temp;
				temp->up->down = nullptr;
				temp = temp->right;
				delete delete_cell;
			}
		}
		else
		{
			c_row--;
			current = current->up;
			while (temp != nullptr)
			{
				delete_cell = temp;
				temp->down->up = temp->up;
				temp->up->down = temp->down;

				temp = temp->right;
				delete delete_cell;
			}
		}
		row_size--;
	}

	void Clear_Col()
	{
		Cell* temp = current;

		while (temp->up != nullptr)
		{
			temp = temp->up;
		}

		while (temp != nullptr)
		{
			temp->data="    ";
			temp = temp->down;
		}

	}

	void Clear_Row()
	{
		Cell* temp = current;

		while (temp->left != nullptr)
		{
			temp = temp->left;
		}

		while (temp != nullptr)
		{
			temp->data="    ";
			temp = temp->right;
		}

	}

	void Print_Grid()
	{
		for (int ri = 0; ri < row_size; ri++)
		{
			for (int ci = 0; ci < col_size; ci++)
			{
				Print_cell(ri, ci, 7);
			}
		}
	}

	void Print_Data()
	{
		Cell* temp = head;
		for (int ri = 0; ri < row_size; ri++)
		{
			Cell* temp2 = temp;
			for (int ci = 0; ci < col_size; ci++)
			{
				Print_cell_data(ri, ci, temp, 7);
				temp = temp->right;
			}

			temp = temp2->down;
		}
	}

	void movement()
	{
		int max_col = 0, max_row = 0, min_col = INT_MAX, min_row = INT_MAX;
		while (true)
		{
			char c = _getch();
			if (c == 100)									//right(d)
			{
				if (current->right != nullptr)
				{
					current = current->right;
					c_col++;
					Print_cell(c_row, c_col, 4);
				}

			}
			else if (c == 97)							//left(a)
			{

				if (current->left != nullptr)
				{
					current = current->left;
					c_col--;
				}

			}
			else if (c == 115)							//down(s)
			{

				if (current->down != nullptr)
				{
					current = current->down;
					c_row++;
				}

			}
			else if (c == 119)							//up(w)
			{
				if (current->up != nullptr)
				{
					current = current->up;
					c_row--;
				}

			}
			else if (c == 99)							//calculate(c)
				break;
			if (c_col > max_col)
				max_col = c_col;
			if (c_row > max_row)
				max_row = c_row;
			if (c_col < min_col)
				min_col = c_col;
			if (c_row < min_row)
				min_row = c_row;

			Print_cell(c_row, c_col, 4);
			Print_cell_data(c_row, c_col, current, 7);
		}

		Range_end.col = max_col;
		Range_end.row = max_row;

		Range_start.col = min_col;
		Range_start.row = min_row;
	}

	bool check_string_digit(Cell* temp)
	{
		for (int i = 0; i < temp->data.length(); i++)
		{
			if (!isdigit(temp->data[i]))
				return false;
		}
		return true;
	}

	int Calculate_sum()
	{
		Cell* temp = RangeStart;
		int sum = 0;
		int sri = c_row;
		int sci = c_col;
		movement();

		int row_limit = Range_end.row - Range_start.row;
		int col_limit = Range_end.col - Range_start.col;


		if (Range_start.col <= sri && Range_start.row < sci)
		{
			for (int i = 0; i <= Range_start.col; i++)
			{
				temp = temp->left;
			}
		}

		for (int ri = 0; ri <= row_limit; ri++)
		{
			Cell* temp2 = temp;
			for (int ci = 0; ci <= col_limit; ci++)
			{
				if (check_string_digit(temp))
				{
					sum = sum + std::stoi(temp->data);

				}
				temp = temp->right;
			}

			temp = temp2->down;

		}

		c_row = sri;
		c_col = sci;
		return sum;
	}

	int Calculate_average()
	{
		Cell* temp = RangeStart;
		int average = 0;
		int sri = c_row;
		int sci = c_col;
		movement();

		int row_limit = Range_end.row - Range_start.row;
		int col_limit = Range_end.col - Range_start.col;


		if (Range_start.col <= sri && Range_start.row < sci)
		{
			for (int i = 0; i <= Range_start.col; i++)
			{
				temp = temp->left;
			}
		}

		for (int ri = 0; ri <= row_limit; ri++)
		{
			Cell* temp2 = temp;
			for (int ci = 0; ci <= col_limit; ci++)
			{
				if (check_string_digit(temp))
				{
					average = average + std::stoi(temp->data);

				}
				temp = temp->right;
			}

			temp = temp2->down;

		}

		c_row = sri;
		c_col = sci;
		return (average / 2);
	}

	int Calculate_Count()
	{
		Cell* temp = RangeStart;
		int count = 0;
		int sri = c_row;
		int sci = c_col;
		movement();

		int row_limit = Range_end.row - Range_start.row;
		int col_limit = Range_end.col - Range_start.col;


		if (Range_start.col <= sri && Range_start.row < sci)
		{
			for (int i = 0; i <= Range_start.col; i++)
			{
				temp = temp->left;
			}
		}

		for (int ri = 0; ri <= row_limit; ri++)
		{
			Cell* temp2 = temp;
			for (int ci = 0; ci <= col_limit; ci++)
			{
				if (temp->data != " ")
					count++;
				temp = temp->right;
			}

			temp = temp2->down;

		}

		c_row = sri;
		c_col = sci;
		return count;
	}

	int Calculate_Max()
	{
		Cell* temp = RangeStart;
		int Max = INT_MIN;
		int sri = c_row;
		int sci = c_col;
		movement();

		int row_limit = Range_end.row - Range_start.row;
		int col_limit = Range_end.col - Range_start.col;


		if (Range_start.col <= sri && Range_start.row < sci)
		{
			for (int i = 0; i <= Range_start.col; i++)
			{
				temp = temp->left;
			}
		}

		for (int ri = 0; ri <= row_limit; ri++)
		{
			Cell* temp2 = temp;
			for (int ci = 0; ci <= col_limit; ci++)
			{
				if (check_string_digit(temp))
				{

					if (Max < std::stoi(temp->data))
					{
						Max = std::stoi(temp->data);
					}

				}
				temp = temp->right;
			}

			temp = temp2->down;

		}

		c_row = sri;
		c_col = sci;
		return Max;
	}

	int Calculate_Min()
	{
		Cell* temp = RangeStart;
		int Min = INT_MAX;
		int sri = c_row;
		int sci = c_col;
		movement();

		int row_limit = Range_end.row - Range_start.row;
		int col_limit = Range_end.col - Range_start.col;


		if (Range_start.col <= sri && Range_start.row < sci)
		{
			for (int i = 0; i <= Range_start.col; i++)
			{
				temp = temp->left;
			}
		}

		for (int ri = 0; ri <= row_limit; ri++)
		{
			Cell* temp2 = temp;
			for (int ci = 0; ci <= col_limit; ci++)
			{
				if (check_string_digit(temp))
				{

					if (Min > std::stoi(temp->data))
					{
						Min = std::stoi(temp->data);
					}

				}
				temp = temp->right;
			}

			temp = temp2->down;

		}

		c_row = sri;
		c_col = sci;
		return Min;
	}

	void Copy()
	{
		Cell* temp = RangeStart;
		Clipboard.clear();
		int sri = c_row;
		int sci = c_col;
		movement();

		int row_limit = Range_end.row - Range_start.row;
		int col_limit = Range_end.col - Range_start.col;


		if (Range_start.col <= sri && Range_start.row < sci)
		{
			for (int i = 0; i <= Range_start.col; i++)
			{
				temp = temp->left;
			}
		}

		for (int ri = 0; ri <= row_limit; ri++)
		{
			vector<string> clip;
			Cell* temp2 = temp;
			for (int ci = 0; ci <= col_limit; ci++)
			{
				clip.push_back(temp->data);
				temp = temp->right;
			}

			Clipboard.push_back(clip);
			temp = temp2->down;

		}

		c_row = sri;
		c_col = sci;
	}

	void Cut()
	{
		Cell* temp = RangeStart;
		Clipboard.clear();
		int sri = c_row;
		int sci = c_col;
		movement();

		int row_limit = Range_end.row - Range_start.row;
		int col_limit = Range_end.col - Range_start.col;


		if (Range_start.col <= sri && Range_start.row < sci)
		{
			for (int i = 0; i <= Range_start.col; i++)
			{
				temp = temp->left;
			}
		}

		for (int ri = 0; ri <= row_limit; ri++)
		{
			vector<string> clip;
			Cell* temp2 = temp;
			for (int ci = 0; ci <= col_limit; ci++)
			{
				clip.push_back(temp->data);
				temp->data = " ";
				temp = temp->right;
			}

			Clipboard.push_back(clip);
			temp = temp2->down;

		}

		c_row = sri;
		c_col = sci;
	}

	void Paste()
	{
		Cell* temp = current;
		for (int ri = 0; ri < Clipboard.size(); ri++)
		{
			Cell* temp2 = current;
			for (int ci = 0; ci < Clipboard[0].size(); ci++)
			{
				current->data = Clipboard[ri][ci];
				if (current->right == nullptr)
					InsertColRight();
				current = current->right;
			}

			if (temp2->down == nullptr)
				InsertRowBelow();
			current = temp2->down;

		}

		current = temp;

	}

	void Mathematic_operations()
	{
		string s;
		char c = _getch();
		RangeStart = current;
		if (c == 115)			//Sum(S)
		{
			s = std::to_string(Calculate_sum());
			color(7);
			gotoRowCol(Cell_height * row_size + 10, 0);
			cout << "Sum = " << s << " ";
		}
		else if (c == 97)			//Average(A)
		{
			s = std::to_string(Calculate_average());
			color(7);
			gotoRowCol(Cell_height * row_size + 10, 0);
			cout << "Average = " << s << " ";

		}
		else if (c == 99)				// Count(c)
		{
			s = std::to_string(Calculate_Count());
			color(7);
			gotoRowCol(Cell_height * row_size + 10, 0);
			cout << "Count = " << s << " ";
		}
		else if (c == 109)				// Max(m)
		{
			s = std::to_string(Calculate_Max());
			color(7);
			gotoRowCol(Cell_height * row_size + 10, 0);
			cout << "Max = " << s << " ";
		}
		else if (c == 110)				// Min(n)
		{
			s = std::to_string(Calculate_Min());
			color(7);
			gotoRowCol(Cell_height * row_size + 10, 0);
			cout << "Min = " << s << " ";
		}
		system("pause");
		gotoRowCol(Cell_height * row_size + 10, 0);
		cout << "                                                          ";
		current = RangeStart;
		current->data = s;
		Print_Grid();
		Print_Data();
	}

	void Shift_and_delete()
	{
		char c = _getch();
		if (c == 100)					//right cell shift(d)
		{

			InsertCellByRightShift();
			Print_Grid();
			Print_Data();
		}
		else if (c == 97)					//Delete shift left(a)
		{
			DeleteCellByLeftShift();
			Print_Grid();
			Print_Data();
		}
		else if (c == 119)					//Delete shift up(w)
		{
			DeleteCellByUpShift();
			Print_Grid();
			Print_Data();
		}
		else if (c == 115)					//down cell shift(s)
		{
			InsertCellByDownShift();
			Print_Grid();
			Print_Data();
		}
		else if (c == 114)				//delete (r)
		{
			Delete_Row();
			system("cls");
			Print_Grid();
			Print_Data();
		}
		else if (c == 99)				//delete (c)
		{
			Delete_Col();
			system("cls");
			Print_Grid();
			Print_Data();
		}
	}

	void Cut_Copy_Paste()
	{
		string s;
		char c = _getch();
		RangeStart = current;
		if (c == 99)			//Copy(C)
		{
			Copy();
		}
		else if (c == 112)				//Paste(P)
		{
			Paste();
		}
		else if (c == 120)				//Cut(x)
		{
			Cut();
		}
		current = RangeStart;
		Print_Grid();
		Print_Data();
	}

	void Keyboard()
	{
		Print_cell(c_row, c_col, 4);
		Cell* temp = current;
		string input;
		while (true)
		{
			char c = _getch();
			if (c == 100)									//right(d)
			{
				if (current->right == nullptr)
				{
					InsertColRight();
					Print_Col();
				}
				Print_cell(c_row, c_col, 7);						//Unhighlight
				Print_cell_data(c_row, c_col, current, 7);
				current = current->right;
				c_col++;
			}
			else if (c == 97)							//left(a)
			{
				if (current->left == nullptr)
				{
					InsertColLeft();
					Print_Col();
					c_col++;
				}
				Print_cell(c_row, c_col, 7);						//Unhighlight
				Print_cell_data(c_row, c_col, current, 7);
				current = current->left;
				c_col--;

			}
			else if (c == 115)							//down(s)
			{
				if (current->down == nullptr)
				{
					InsertRowBelow();;
					Print_Row();

				}
				Print_cell(c_row, c_col, 7);						//Unhighlight
				Print_cell_data(c_row, c_col, current, 7);
				current = current->down;
				c_row++;
			}
			else if (c == 119)							//up(w)
			{
				if (current->up == nullptr)
				{
					InsertRowAbove();
					Print_Row();
					c_row++;
				}
				Print_cell(c_row, c_col, 7);						//Unhighlight
				Print_cell_data(c_row, c_col, current, 7);
				current = current->up;
				c_row--;

			}
			else if (c == 105)							//insertion(i)
			{
				do
				{
					gotoRowCol(Cell_height * row_size + 10, 0);
					Print_cell_data(c_row, c_col, current, 7);
					gotoRowCol(Cell_height * row_size + 10, 0);
					cout << "Enter the value:";
					cin >> input;

					if (input.length() > 4)
					{
						gotoRowCol(Cell_height * row_size + 10 + 1, 0);
						cout << "Invalid input";
						gotoRowCol(Cell_height * row_size + 10, 0);
						cout << "                                                                    ";
					}
					else
					{
						gotoRowCol(Cell_height * row_size + 10 + 1, 0);
						cout << "                                                                     ";
						gotoRowCol(Cell_height * row_size + 10, 0);
						cout << "                                                                    ";
						break;
					}

				} while (true);
				current->data = input;
				Print_cell_data(c_row, c_col, current, 7);

			}

			else if (c == 111)					//Open Mneu(o)
			{
				Shift_and_delete();
			}

			else if (c == 99)			//Clear row and col(c)
			{
				c = _getch();
				if (c == 99)
				{
					Clear_Col();
					Print_Data();			//Clear column(c)
				}
				else if (c == 114)
				{
					Clear_Row();			//Clear row(r)
					Print_Data();
				}
			}

			else if (c == 109)			//Mathematic operations(m)
			{
				Mathematic_operations();
			}

			else if (c == 49)				//Save(1)
			{
				save_file();
			}
			else if (c == 50)
			{								//Load(2)
				load_file();
			}

			else if (c == 120)				//Cut and Copy(x)
			{
				Cut_Copy_Paste();
			}
			else if (c == 48)			//Menu(0)
			{
				menu();
			}
			gotoRowCol(c_row, c_col);
			Print_cell(c_row, c_col, 4);						//Highlight new cell
			Print_cell_data(c_row, c_col, current, 7);
			
		}

	}

	void save_file()
	{
		Cell* temp = head;
		ofstream fout("Save.txt");
		fout << row_size << endl;
		fout << col_size << endl;
		for (int i = 0; i < row_size; i++)
		{
			Cell* temp2 = temp;
			for (int j = 0; j < col_size; j++)
			{
				if (temp->data == " ")
				{
					fout << "Space" << " ";
				}
				else
				{
					fout << temp->data << " ";
				}

				temp = temp->right;
			}
			fout << endl;
			temp = temp2->down;
		}
	}

	void load_file()
	{
		system("cls");
		ifstream fin("Save.txt");
		fin >> row_size;
		fin >> col_size;

		head = nullptr;
		current = nullptr;
		c_row = c_col = 0;

		head = New_row();
		current = head;
		for (int i = 0; i < row_size - 1; i++)
		{
			InsertRowBelow();
			row_size--;
			current = current->down;
		}

		string data;
		current = head;
		Cell* temp = current;
		for (int ri = 0; ri < row_size; ri++)
		{
			Cell* temp2 = temp;
			for (int ci = 0; ci < col_size; ci++)
			{
				fin >> data;
				if (data == "Space")
					temp->data = " ";
				else
					temp->data = data;

				temp = temp->right;
			}

			temp = temp2->down;
		}

		Print_Grid();
		Print_Data();
	}

	void menu()
	{

		color(7);
		int x = 15;   //Spacing for column
		gotoRowCol(0, col_size * x);
		cout << "Shortcut Keys: ";
		gotoRowCol(1, col_size * x);
		cout << "Insert : i";
		gotoRowCol(2, col_size * x);
		cout << "Shift Cell Right O -> D";
		gotoRowCol(3, col_size * x);
		cout << "Shift Cell Left O -> A";
		gotoRowCol(4, col_size * x);
		cout << "Shift Cell Up O -> U";
		gotoRowCol(5, col_size * x);
		cout << "Shift Cell Down O -> S";
		gotoRowCol(6, col_size * x);
		cout << "Delete Row O -> R";
		gotoRowCol(7, col_size * x);
		cout << "Delete Col O -> C";
		gotoRowCol(8, col_size * x);
		cout << "Copy X -> C and C to select";
		gotoRowCol(9, col_size * x);
		cout << "Cut X -> X and C";
		gotoRowCol(10, col_size * x);
		cout << "Paste X -> P and C";
		gotoRowCol(11, col_size * x);
		cout << "Clear Row C -> C";
		gotoRowCol(12, col_size * x);
		cout << "Clear Col C -> R";
		gotoRowCol(13, col_size * x);
		cout << "Save File 1";
		gotoRowCol(14, col_size * x);
		cout << "Load File 2";
		gotoRowCol(15, col_size * x);
		cout << "Sum Range M -> S and C to select";
		gotoRowCol(16, col_size * x);
		cout << "Average Range M -> A and C";
		gotoRowCol(17, col_size * x);
		cout << "Count Range M -> C and C";
		gotoRowCol(18, col_size * x);
		cout << "Max in Range M -> M and C";
		gotoRowCol(19, col_size * x);
		cout << "Min in Range M -> n and C" ;


		/*color(0);*/
		gotoRowCol(20, col_size * x);
		system("pause");
		gotoRowCol(0, col_size * x);
		cout << "                                        ";
		gotoRowCol(1, col_size * x);
		cout << "                                        ";
		gotoRowCol(2, col_size * x);
		cout << "                                        ";
		gotoRowCol(3, col_size * x);
		cout << "                                        ";
		gotoRowCol(4, col_size * x);
		cout << "                                        ";
		gotoRowCol(5, col_size * x);
		cout << "                                        ";
		gotoRowCol(6, col_size * x);
		cout << "                                        ";
		gotoRowCol(7, col_size * x);
		cout << "                                        ";
		gotoRowCol(8, col_size * x);
		cout << "                                        ";
		gotoRowCol(9, col_size * x);
		cout << "                                        ";
		gotoRowCol(10, col_size * x);
		cout << "                                        ";
		gotoRowCol(11, col_size * x);
		cout << "                                        ";
		gotoRowCol(12, col_size * x);
		cout << "                                        ";
		gotoRowCol(13, col_size * x);
		cout << "                                        ";
		gotoRowCol(14, col_size * x);
		cout << "                                        ";
		gotoRowCol(15, col_size * x);
		cout << "                                        ";
		gotoRowCol(16, col_size * x);
		cout << "                                        ";
		gotoRowCol(17, col_size * x);
		cout << "                                        ";
		gotoRowCol(18, col_size * x);
		cout << "                                        ";
		gotoRowCol(19, col_size * x);
		cout << "                                        ";
		gotoRowCol(20, col_size * x);
		cout << "                                        ";
	
	}

	
};



int main()
{
	Excel data;
	data.Keyboard();


}

