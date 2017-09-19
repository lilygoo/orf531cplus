#include <xlw/xlw.h>
using namespace xlw;

// Registration of Excel functions
namespace 
{
  // Register the function SAYHELLO.
  XLRegistration::Arg SayHelloArgs[] = {
    { "Name", "You name", "XLF_OPER" }
  };
  XLRegistration::XLFunctionRegistrationHelper registerSayHello(
    "xlSayHello", "SAYHELLO", "Says Hello",
    "XlwTest", SayHelloArgs, 1);

  // Register the function OUTERPROD.
  XLRegistration::Arg OuterProdArgs[] = {
    { "Vec1", "The first vector", "XLF_OPER" },
    { "Vec2", "The second vector", "XLF_OPER" }
  };
  XLRegistration::XLFunctionRegistrationHelper registerOuterProd(
    "xlOuterProd", "OUTERPROD", "Computes the outer product of two numeric vectors",
    "XlwTest", OuterProdArgs, 2);

} // namespace
