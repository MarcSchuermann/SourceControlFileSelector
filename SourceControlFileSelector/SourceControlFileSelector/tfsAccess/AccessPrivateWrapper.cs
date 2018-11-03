//// --------------------------------------------------------------------------------------------------------------------
//// <copyright>Marc Schürmann</copyright>
//// --------------------------------------------------------------------------------------------------------------------

using System.Dynamic;
using System.Linq;
using System.Reflection;

namespace SourceControlFileSelector
{
    /// <summary>Access to privates.</summary>
    /// <seealso cref="System.Dynamic.DynamicObject"/>
    public class AccessPrivateWrapper : DynamicObject
    {
        #region Private Fields

        /// <summary>The flags</summary>
        private static BindingFlags flags = BindingFlags.NonPublic | BindingFlags.Instance
            | BindingFlags.Static | BindingFlags.Public;

        private object wrappedObject;

        #endregion Private Fields

        #region Public Constructors

        /// <summary>Initializes a new instance of the <see cref="AccessPrivateWrapper"/> class.</summary>
        /// <param name="objectToWrapp">The object to wrapp.</param>
        public AccessPrivateWrapper(object objectToWrapp)
        {
            wrappedObject = objectToWrapp;
        }

        #endregion Public Constructors

        #region Public Methods

        /// <summary>Froms the type.</summary>
        /// <param name="asm">The asm.</param>
        /// <param name="type">The type.</param>
        /// <param name="args">The arguments.</param>
        /// <returns></returns>
        public static dynamic FromType(Assembly asm, string type, params object[] args)
        {
            var allt = asm.GetTypes();
            var t = allt.First(item => item.Name == type);

            var types = from a in args
                        select a.GetType();

            var ctor = t.GetConstructor(flags, null, types.ToArray(), null);

            if (ctor != null)
            {
                var instance = ctor.Invoke(args);
                return new AccessPrivateWrapper(instance);
            }

            return null;
        }

        /// <summary>
        /// Provides the implementation for operations that get member values. Classes derived from
        /// the <see cref="T:System.Dynamic.DynamicObject"/> class can override this method to
        /// specify dynamic behavior for operations such as getting a value for a property.
        /// </summary>
        /// <param name="binder">
        /// Provides information about the object that called the dynamic operation. The binder.Name
        /// property provides the name of the member on which the dynamic operation is performed. For
        /// example, for the Console.WriteLine(sampleObject.SampleProperty) statement, where
        /// sampleObject is an instance of the class derived from the
        /// <see cref="T:System.Dynamic.DynamicObject"/> class, binder.Name returns "SampleProperty".
        /// The binder.IgnoreCase property specifies whether the member name is case-sensitive.
        /// </param>
        /// <param name="result">
        /// The result of the get operation. For example, if the method is called for a property, you
        /// can assign the property value to <paramref name="result"/>.
        /// </param>
        /// <returns>
        /// true if the operation is successful; otherwise, false. If this method returns false, the
        /// run-time binder of the language determines the behavior. (In most cases, a run-time
        /// exception is thrown.)
        /// </returns>
        public override bool TryGetMember(GetMemberBinder binder, out object result)
        {
            var prop = wrappedObject.GetType().GetProperty(binder.Name, flags);

            if (prop == null)
            {
                var fld = wrappedObject.GetType().GetField(binder.Name, flags);
                if (fld != null)
                {
                    result = fld.GetValue(wrappedObject);
                    return true;
                }
                else
                    return base.TryGetMember(binder, out result);
            }
            else
            {
                result = prop.GetValue(wrappedObject, null);
                return true;
            }
        }

        /// <summary>
        /// Provides the implementation for operations that invoke a member. Classes derived from the
        /// <see cref="T:System.Dynamic.DynamicObject"/> class can override this method to specify
        /// dynamic behavior for operations such as calling a method.
        /// </summary>
        /// <param name="binder">
        /// Provides information about the dynamic operation. The binder.Name property provides the
        /// name of the member on which the dynamic operation is performed. For example, for the
        /// statement sampleObject.SampleMethod(100), where sampleObject is an instance of the class
        /// derived from the <see cref="T:System.Dynamic.DynamicObject"/> class, binder.Name returns
        /// "SampleMethod". The binder.IgnoreCase property specifies whether the member name is case-sensitive.
        /// </param>
        /// <param name="args">
        /// The arguments that are passed to the object member during the invoke operation. For
        /// example, for the statement sampleObject.SampleMethod(100), where sampleObject is derived
        /// from the <see cref="T:System.Dynamic.DynamicObject"/> class, <paramref name="args[0]"/>
        /// is equal to 100.
        /// </param>
        /// <param name="result">The result of the member invocation.</param>
        /// <returns>
        /// true if the operation is successful; otherwise, false. If this method returns false, the
        /// run-time binder of the language determines the behavior. (In most cases, a
        /// language-specific run-time exception is thrown.)
        /// </returns>
        public override bool TryInvokeMember(InvokeMemberBinder binder, object[] args, out object result)
        {
            var types = from a in args
                        select a.GetType();

            var method = wrappedObject.GetType().GetMethod
                (binder.Name, flags, null, types.ToArray(), null);

            if (method == null)
                return base.TryInvokeMember(binder, args, out result);
            else
            {
                result = method.Invoke(wrappedObject, args);
                return true;
            }
        }

        /// <summary>
        /// Provides the implementation for operations that set member values. Classes derived from
        /// the <see cref="T:System.Dynamic.DynamicObject"/> class can override this method to
        /// specify dynamic behavior for operations such as setting a value for a property.
        /// </summary>
        /// <param name="binder">
        /// Provides information about the object that called the dynamic operation. The binder.Name
        /// property provides the name of the member to which the value is being assigned. For
        /// example, for the statement sampleObject.SampleProperty = "Test", where sampleObject is an
        /// instance of the class derived from the <see cref="T:System.Dynamic.DynamicObject"/>
        /// class, binder.Name returns "SampleProperty". The binder.IgnoreCase property specifies
        /// whether the member name is case-sensitive.
        /// </param>
        /// <param name="value">
        /// The value to set to the member. For example, for sampleObject.SampleProperty = "Test",
        /// where sampleObject is an instance of the class derived from the
        /// <see cref="T:System.Dynamic.DynamicObject"/> class, the <paramref name="value"/> is "Test".
        /// </param>
        /// <returns>
        /// true if the operation is successful; otherwise, false. If this method returns false, the
        /// run-time binder of the language determines the behavior. (In most cases, a
        /// language-specific run-time exception is thrown.)
        /// </returns>
        public override bool TrySetMember(SetMemberBinder binder, object value)
        {
            var prop = wrappedObject.GetType().GetProperty(binder.Name, flags);
            if (prop == null)
            {
                var fld = wrappedObject.GetType().GetField(binder.Name, flags);
                if (fld != null)
                {
                    fld.SetValue(wrappedObject, value);
                    return true;
                }
                else
                    return base.TrySetMember(binder, value);
            }
            else
            {
                prop.SetValue(wrappedObject, value, null);
                return true;
            }
        }

        #endregion Public Methods
    }
}