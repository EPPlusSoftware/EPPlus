using System;
using System.Linq.Expressions;

namespace OfficeOpenXml.DataValidation
{
    public static class New<T>
        where T : new()
    {
        public static Func<T> Instance =
            Expression.Lambda<Func<T>>(Expression.New(typeof(T))).Compile();
    }

    public static class CompiledLambdaExpressions
    {

        public static T CreateInstance<T>(this Type type)
            where T : new() => New<T>.Instance();

        //public static object CreateInstance<TArg>(this Type type, TArg arg)
        //    => CreateInstance<TArg, TypeToIgnore>(type, arg, null);

        //public static object CreateInstance<TArg1, TArg2>(this Type type, TArg1 arg1, TArg2 arg2)
        //    => CreateInstance<TArg1, TArg2, TypeToIgnore>(type, arg1, arg2, null);

        private class TypeToIgnore
        {

        }

        //private readonly Container _container;

        //internal static void CreateLambda(string uid, string address, ExcelDataValidationType validationType)
        //{

        //    switch (validationType.Type)
        //    {
        //        case eDataValidationType.Any:
        //            return new NewDataValidationAny(uid, address, validationType);
        //        default:
        //            throw new InvalidOperationException("Non supported validationtype: " + validationType.Type.ToString());
        //    }
        //}
    }
}
