// DynArray2D.h
// Code from https://www.qtcentre.org/threads/31440-two-dimensional-array-size-determined-dynamically
// Some code is fixed by j2doll

#ifndef DYNARRAY2D_H
#define DYNARRAY2D_H

template <typename T> class DynArray2D
{
public:
    DynArray2D(unsigned int n, unsigned int m)
    {
        _n = n;
        _m = m;

        _array = new T*[n];

        for(unsigned  int i = 0; i < n; i++)
        {
            _array[i] = new T[m];
        }
    }

    void setValue(unsigned int n, unsigned int m, T v)
    {
        _array[n][m] = v;
    }

    T getValue(unsigned int n, unsigned int m)
    {
        return _array[n][m];
    }

    ~DynArray2D()
    {
        for (unsigned int i = 0 ; i < _n ; i++)
        {
            delete [] _array[i];
        }
        delete [] _array;
    }

protected:
    T **_array;
    unsigned int _n;
    unsigned int _m;
};

#endif // DYNARRAY2D_H
