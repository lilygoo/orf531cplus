#include <xlw/xlw.h>
#include <vector>
using namespace xlw;
using namespace std;

extern "C" {

LPXLFOPER EXCEL_EXPORT xlSayHello(LPXLFOPER xlName) {
EXCEL_BEGIN;

  if (XlfExcel::Instance().IsCalledByFuncWiz())
    return XlfOper(true);

  string name = XlfOper(xlName).AsString();       // read in the name
  string greet = "Hello " + name + "!";  // compose the greeting
  return XlfOper(greet);                 // return as XlfOper
EXCEL_END;
}

LPXLFOPER EXCEL_EXPORT xlOuterProd(LPXLFOPER xlVec1, LPXLFOPER xlVec2) {
EXCEL_BEGIN;

  if (XlfExcel::Instance().IsCalledByFuncWiz())
    return XlfOper(true);

  // read the input vectors
  vector<double> v1 = XlfOper(xlVec1).AsDoubleVector();
  vector<double> v2 = XlfOper(xlVec2).AsDoubleVector();
  // allocate the returned matrix
  RW n1 = v1.size();
  COL n2 = v2.size();
  XlfOper ret(n1, n2);
  // calculate the outer product and store the result
  for (RW i = 0; i < n1; ++i) {
      double tmp = v1[i]; 
      for (COL j = 0; j < n2; ++j) {
          ret.SetElement(i, j, tmp * v2[j]);
      }
  }
  return ret;
EXCEL_END;
}

} // extern "C"
