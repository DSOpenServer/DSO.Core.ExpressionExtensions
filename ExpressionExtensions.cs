using System.Linq.Expressions;
using System.Reflection;

namespace DSO.Core.ExpressionExtensions
{
    public static class ExpressionExtensions
    {
        #region DSO
        private static readonly object _ToDelegateArrayLock = new object();
        private static readonly object _ToDictionaryDelegateLock = new object();

        private static readonly Dictionary<Type, Delegate> _cachedToDictionaryDelegate = new Dictionary<Type, Delegate>();
        private static readonly Dictionary<Type, (Delegate Factory, Type[] ColumnTypes, string[] ColumnNames)> _cachedToDelegateArray = new Dictionary<Type, (Delegate Factory, Type[] ColumnTypes, string[] ColumnNames)>();

        public static Delegate[] ToDelegateArray<T>(this T source, out Type[] ColumnTypes, out string[] ColumnNames) where T : class
        {
            if (source == null)
            {
                ColumnNames = null;
                ColumnTypes = null;
                return null;
            }

            var sourceType = typeof(T);

            Func<T, Delegate[]> func = null;
            lock (_ToDelegateArrayLock)
            {
                if (_cachedToDelegateArray.TryGetValue(sourceType, out var cacheEntry))
                {
                    func = (Func<T, Delegate[]>)cacheEntry.Factory;
                    ColumnNames = cacheEntry.ColumnNames;
                    ColumnTypes = cacheEntry.ColumnTypes;
                }
                else
                {
                    func = CreateDelegateArrayFactory<T>(out ColumnTypes, out ColumnNames);
                    _cachedToDelegateArray.Add(sourceType, (func, ColumnTypes, ColumnNames));
                }
            }

            return func(source);
        }

        private static Func<T, Delegate[]> CreateDelegateArrayFactory<T>(out Type[] ColumnTypes, out string[] ColumnNames) where T : class
        {
            var sourceType = typeof(T);
            var sourceParam = Expression.Parameter(sourceType, "source");
            var array = Expression.Variable(typeof(Delegate[]), "array");

            var members = sourceType.GetMembers(BindingFlags.Public | BindingFlags.Instance)
                .Where(x => x.MemberType == MemberTypes.Property || x.MemberType == MemberTypes.Field);
            var picked = members.Where(m => (m is PropertyInfo pp && pp.CanRead && pp.GetIndexParameters().Length == 0) || m is FieldInfo).ToArray();

            var expressions = new List<Expression>();
            var arrayNew = Expression.NewArrayBounds(typeof(Delegate), Expression.Constant(picked.Length));
            expressions.Add(Expression.Assign(array, arrayNew));
            ColumnNames = new string[picked.Length];
            ColumnTypes = new Type[picked.Length];

            for (int i = 0; i < picked.Length; i++)
            {
                var member = picked[i];
                ColumnNames[i] = member.Name;

                Type memberType = member.MemberType == MemberTypes.Property
                    ? ((PropertyInfo)member).PropertyType
                    : ((FieldInfo)member).FieldType;

                ColumnTypes[i] = memberType;

                var memberAccess = Expression.PropertyOrField(sourceParam, member.Name);
                var funcType = typeof(Func<>).MakeGenericType(memberType);
                var lambda = Expression.Lambda(funcType, memberAccess);

                expressions.Add(Expression.Assign(
                    Expression.ArrayAccess(array, Expression.Constant(i)),
                    lambda
                ));
            }

            expressions.Add(array);
            var block = Expression.Block(new[] { array }, expressions);
            return Expression.Lambda<Func<T, Delegate[]>>(block, sourceParam).Compile();
        }

        public static Dictionary<string, Delegate> ToDictionaryDelegate<T>(this T source) where T : class
        {
            if (source == null) return null;

            var sourceType = typeof(T);

            Func<T, Dictionary<string, Delegate>> func = null;
            lock (_ToDictionaryDelegateLock)
            {
                if (_cachedToDictionaryDelegate.TryGetValue(sourceType, out Delegate del))
                {
                    func = (Func<T, Dictionary<string, Delegate>>)del;
                }
                else
                {
                    func = CreateToDictionaryDelegateFactory<T>();
                    _cachedToDictionaryDelegate.Add(sourceType, func);
                }
            }

            return func(source);
        }

        private static Func<T, Dictionary<string, Delegate>> CreateToDictionaryDelegateFactory<T>() where T : class
        {
            var sourceType = typeof(T);
            var sourceParam = Expression.Parameter(sourceType, "source");
            var dict = Expression.Variable(typeof(Dictionary<string, Delegate>), "dict");
            var addMethod = typeof(Dictionary<string, Delegate>).GetMethod("Add");

            var members = sourceType.GetMembers(BindingFlags.Public | BindingFlags.Instance).Where(x => x.MemberType == MemberTypes.Property || x.MemberType == MemberTypes.Field);

            var picked = members.Where(m => (m is PropertyInfo pp && pp.CanRead && pp.GetIndexParameters().Length == 0) || m is FieldInfo).ToArray();
            var newDictCtor = typeof(Dictionary<string, Delegate>).GetConstructor(new[] { typeof(int) });

            var expressions = new List<Expression>();

            var dictNew = Expression.New(newDictCtor,
                Expression.Constant(picked.Length));

            expressions.Add(Expression.Assign(dict, dictNew));

            foreach (var member in picked)
            {
                Type memberType = member.MemberType == MemberTypes.Property
                    ? ((PropertyInfo)member).PropertyType
                    : ((FieldInfo)member).FieldType;

                // Üye erişim ifadesi
                var memberAccess = Expression.PropertyOrField(sourceParam, member.Name);

                // Func<TProperty> tipi için lambda oluştur
                var funcType = typeof(Func<>).MakeGenericType(memberType);
                // Lambda ifadesi: () => source.Member
                var lambda = Expression.Lambda(funcType, memberAccess);

                // Dictionary'e ekle
                expressions.Add(Expression.Call(
                    dict,
                    addMethod,
                    Expression.Constant(member.Name),
                    lambda
                ));
            }

            expressions.Add(dict);
            var block = Expression.Block(new[] { dict }, expressions);
            return Expression.Lambda<Func<T, Dictionary<string, Delegate>>>(block, sourceParam).Compile();
        }
        #endregion DSO
    }
}
