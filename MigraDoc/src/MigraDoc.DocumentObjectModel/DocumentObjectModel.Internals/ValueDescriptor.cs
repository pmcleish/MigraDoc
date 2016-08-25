#region MigraDoc - Creating Documents on the Fly
//
// Authors:
//   Stefan Lange
//   Klaus Potzesny
//   David Stephensen
//
// Copyright (c) 2001-2016 empira Software GmbH, Cologne Area (Germany)
//
// http://www.pdfsharp.com
// http://www.migradoc.com
// http://sourceforge.net/projects/pdfsharp
//
// Permission is hereby granted, free of charge, to any person obtaining a
// copy of this software and associated documentation files (the "Software"),
// to deal in the Software without restriction, including without limitation
// the rights to use, copy, modify, merge, publish, distribute, sublicense,
// and/or sell copies of the Software, and to permit persons to whom the
// Software is furnished to do so, subject to the following conditions:
//
// The above copyright notice and this permission notice shall be included
// in all copies or substantial portions of the Software.
//
// THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
// IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
// FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL
// THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
// LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING
// FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER 
// DEALINGS IN THE SOFTWARE.
#endregion

using System;
using System.Diagnostics;
using System.Reflection;
using System.Reflection.Emit;

#pragma warning disable 1591

namespace MigraDoc.DocumentObjectModel.Internals
{
    /// <summary>
    /// Base class of all value descriptor classes.
    /// </summary>
    public abstract class ValueDescriptor
    {
        private ValueGetter _valueGetter;
        private ValueSetter _valueSetter;

        private delegate object ValueGetter(DocumentObject dom);
        private delegate void ValueSetter(DocumentObject dom, object value);

        internal ValueDescriptor(string valueName, Type valueType, Type memberType, MemberInfo memberInfo, VDFlags flags)
        {
            // Take new naming convention into account.
            if (valueName.StartsWith("_"))
                valueName = valueName.Substring(1);

            ValueName = valueName;
            ValueType = valueType;
            MemberType = memberType;
            MemberInfo = memberInfo;
            _flags = flags;
        }

        public object CreateValue()
        {
            return Activator.CreateInstance(ValueType, true);
        }

        public abstract object GetValue(DocumentObject dom, GV flags);
        public abstract void SetValue(DocumentObject dom, object val);
        public abstract void SetNull(DocumentObject dom);
        public abstract bool IsNull(DocumentObject dom);

        internal static ValueDescriptor CreateValueDescriptor(MemberInfo memberInfo, DVAttribute attr)
        {
            VDFlags flags = VDFlags.None;
            if (attr.RefOnly)
                flags |= VDFlags.RefOnly;

            string name = memberInfo.Name;

            FieldInfo fieldInfo = memberInfo as FieldInfo;
            Type type = fieldInfo != null ? fieldInfo.FieldType : ((PropertyInfo)memberInfo).PropertyType;

            if (type == typeof(NBool))
                return new NullableDescriptor(name, typeof(Boolean), type, memberInfo, flags);

            if (type == typeof(NInt))
                return new NullableDescriptor(name, typeof(Int32), type, memberInfo, flags);

            if (type == typeof(NDouble))
                return new NullableDescriptor(name, typeof(Double), type, memberInfo, flags);

            if (type == typeof(NString))
                return new NullableDescriptor(name, typeof(String), type, memberInfo, flags);

            if (type == typeof(String))
                return new ValueTypeDescriptor(name, typeof(String), type, memberInfo, flags);

            if (type == typeof(NEnum))
            {
                Type valueType = attr.Type;
#if !NETFX_CORE
                Debug.Assert(valueType.IsSubclassOf(typeof(Enum)), "NEnum must have 'Type' attribute with the underlying type");
#else
                Debug.Assert(valueType.GetTypeInfo().IsSubclassOf(typeof(Enum)), "NEnum must have 'Type' attribute with the underlying type");
#endif
                return new NullableDescriptor(name, valueType, type, memberInfo, flags);
            }

#if !NETFX_CORE
            if (type.IsSubclassOf(typeof(ValueType)))
#else
            if (type.GetTypeInfo().IsSubclassOf(typeof(ValueType)))
#endif
                return new ValueTypeDescriptor(name, type, type, memberInfo, flags);

#if !NETFX_CORE
            if (typeof(DocumentObjectCollection).IsAssignableFrom(type))
#else
            if (typeof(DocumentObjectCollection).GetTypeInfo().IsAssignableFrom(type.GetTypeInfo()))
#endif
                return new DocumentObjectCollectionDescriptor(name, type, type, memberInfo, flags);

#if !NETFX_CORE
            if (typeof(DocumentObject).IsAssignableFrom(type))
#else
            if (typeof(DocumentObject).GetTypeInfo().IsAssignableFrom(type.GetTypeInfo()))
#endif
                return new DocumentObjectDescriptor(name, type, type, memberInfo, flags);

            Debug.Assert(false, type.FullName);
            return null;
        }

        protected void InternalSetValue(DocumentObject dom, object value)
        {
            if (_valueSetter == null)
                _valueSetter = CreateValueSetter();

            _valueSetter(dom, value);
        }

        protected object InternalGetValue(DocumentObject dom)
        {
            if (_valueGetter == null)
                _valueGetter = CreateValueGetter();

            return _valueGetter(dom);
        }

        private ValueSetter CreateValueSetter()
        {
            string methodName = MemberInfo.ReflectedType.FullName + ".set_" + MemberInfo.Name;
            DynamicMethod getterMethod = new DynamicMethod(methodName, null, new[] { typeof(DocumentObject), typeof(object) }, true);
            ILGenerator gen = getterMethod.GetILGenerator();

            gen.Emit(OpCodes.Ldarg_0); // Load DocumentObject
            gen.Emit(OpCodes.Castclass, MemberInfo.ReflectedType);
            gen.Emit(OpCodes.Ldarg_1); // Load value

            FieldInfo field = FieldInfo;

            if (field != null)
            {
                if (field.FieldType.IsValueType)
                    gen.Emit(OpCodes.Unbox_Any, field.FieldType);
                else if (field.FieldType != typeof(object))
                    gen.Emit(OpCodes.Castclass, field.FieldType);

                gen.Emit(OpCodes.Stfld, field);
            }
            else
            {
                PropertyInfo property = PropertyInfo;

                if (property.PropertyType.IsValueType)
                    gen.Emit(OpCodes.Unbox_Any, property.PropertyType);
                else if (property.PropertyType != typeof(object))
                    gen.Emit(OpCodes.Castclass, property.PropertyType);

                gen.Emit(OpCodes.Callvirt, property.GetSetMethod(true));
            }

            gen.Emit(OpCodes.Ret);

            return (ValueSetter)getterMethod.CreateDelegate(typeof(ValueSetter));
        }

        private ValueGetter CreateValueGetter()
        {
            string methodName = MemberInfo.ReflectedType.FullName + ".get_" + MemberInfo.Name;
            DynamicMethod getterMethod = new DynamicMethod(methodName, typeof(object), new[] { typeof(DocumentObject) }, true);
            ILGenerator gen = getterMethod.GetILGenerator();

            gen.Emit(OpCodes.Ldarg_0); // Load DocumentObject
            gen.Emit(OpCodes.Castclass, MemberInfo.ReflectedType);

            FieldInfo field = FieldInfo;

            if (field != null)
            {
                gen.Emit(OpCodes.Ldfld, field);

                if (field.FieldType.IsValueType)
                    gen.Emit(OpCodes.Box, field.FieldType);
            }
            else
            {
                PropertyInfo property = PropertyInfo;

                gen.Emit(OpCodes.Callvirt, property.GetGetMethod(true));

                if (PropertyInfo.PropertyType.IsValueType)
                    gen.Emit(OpCodes.Box, property.PropertyType);
            }

            gen.Emit(OpCodes.Ret);

            return (ValueGetter)getterMethod.CreateDelegate(typeof(ValueGetter));
        }

        public bool IsRefOnly
        {
            get { return (_flags & VDFlags.RefOnly) == VDFlags.RefOnly; }
        }

        public FieldInfo FieldInfo
        {
            get { return MemberInfo as FieldInfo; }
        }

        public PropertyInfo PropertyInfo
        {
            get { return MemberInfo as PropertyInfo; }
        }

        /// <summary>
        /// Name of the value.
        /// </summary>
        public readonly string ValueName;

        /// <summary>
        /// Type of the described value, e.g. typeof(Int32) for an NInt.
        /// </summary>
        public readonly Type ValueType;

        /// <summary>
        /// Type of the described field or property, e.g. typeof(NInt) for an NInt.
        /// </summary>
        public readonly Type MemberType;

        /// <summary>
        /// FieldInfo of the described field.
        /// </summary>
        public readonly MemberInfo MemberInfo;

        /// <summary>
        /// Flags of the described field, e.g. RefOnly.
        /// </summary>
        readonly VDFlags _flags;
    }

    /// <summary>
    /// Value descriptor of all nullable types.
    /// </summary>
    internal class NullableDescriptor : ValueDescriptor
    {
        internal NullableDescriptor(string valueName, Type valueType, Type fieldType, MemberInfo memberInfo, VDFlags flags)
            : base(valueName, valueType, fieldType, memberInfo, flags)
        { }

        public override object GetValue(DocumentObject dom, GV flags)
        {
            Debug.Assert(Enum.IsDefined(typeof(GV), flags), DomSR.InvalidEnumValue(flags));

            object val = InternalGetValue(dom);
            INullableValue ival = (INullableValue)val;
            if (ival.IsNull && flags == GV.GetNull)
                return null;
            return ival.GetValue();
        }

        public override void SetValue(DocumentObject dom, object value)
        {
            object val = InternalGetValue(dom);
            INullableValue ival = (INullableValue)val;

            ival.SetValue(value);
            InternalSetValue(dom, ival);
        }

        public override void SetNull(DocumentObject dom)
        {
            object val = InternalGetValue(dom);
            INullableValue ival = (INullableValue)val;

            ival.SetNull();
            InternalSetValue(dom, ival);
        }

        /// <summary>
        /// Determines whether the given DocumentObject is null (not set).
        /// </summary>
        public override bool IsNull(DocumentObject dom)
        {
            object val = InternalGetValue(dom);

            return ((INullableValue)val).IsNull;
        }
    }

    /// <summary>
    /// Value descriptor of value types.
    /// </summary>
    internal class ValueTypeDescriptor : ValueDescriptor
    {
        internal ValueTypeDescriptor(string valueName, Type valueType, Type fieldType, MemberInfo memberInfo, VDFlags flags)
            : base(valueName, valueType, fieldType, memberInfo, flags)
        { }

        public override object GetValue(DocumentObject dom, GV flags)
        {
            Debug.Assert(Enum.IsDefined(typeof(GV), flags), DomSR.InvalidEnumValue(flags));

            object val = InternalGetValue(dom);
            INullableValue ival = val as INullableValue;
            if (ival != null && ival.IsNull && flags == GV.GetNull)
                return null;
            return val;
        }

        public override void SetValue(DocumentObject dom, object value)
        {
            InternalSetValue(dom, value);
        }

        public override void SetNull(DocumentObject dom)
        {
            object val = InternalGetValue(dom);
            INullableValue ival = (INullableValue)val;

            ival.SetNull();
            InternalSetValue(dom, ival);
        }

        /// <summary>
        /// Determines whether the given DocumentObject is null (not set).
        /// </summary>
        public override bool IsNull(DocumentObject dom)
        {
            object val = InternalGetValue(dom);
            INullableValue ival = val as INullableValue;
            if (ival != null)
                return ival.IsNull;
            return false;
        }
    }

    /// <summary>
    /// Value descriptor of DocumentObject.
    /// </summary>
    internal class DocumentObjectDescriptor : ValueDescriptor
    {
        internal DocumentObjectDescriptor(string valueName, Type valueType, Type fieldType, MemberInfo memberInfo, VDFlags flags)
            : base(valueName, valueType, fieldType, memberInfo, flags)
        { }

        public override object GetValue(DocumentObject dom, GV flags)
        {
            Debug.Assert(Enum.IsDefined(typeof(GV), flags), DomSR.InvalidEnumValue(flags));

            DocumentObject val = InternalGetValue(dom) as DocumentObject;

            if (FieldInfo != null)
            {
                // Member is a field
                if (val == null && flags == GV.ReadWrite)
                {
                    val = (DocumentObject)CreateValue();
                    val._parent = dom;
                    InternalSetValue(dom, val);
                    return val;
                }
            }
            // Member is property
            if (val != null && (val.IsNull() && flags == GV.GetNull))
                return null;

            return val;
        }

        public override void SetValue(DocumentObject dom, object val)
        {
            FieldInfo fieldInfo = FieldInfo;
            // Member is a field
            if (fieldInfo != null)
            {
                InternalSetValue(dom, val);
                return;
            }
            throw new InvalidOperationException("This value cannot be set.");
        }

        public override void SetNull(DocumentObject dom)
        {
            DocumentObject val = InternalGetValue(dom) as DocumentObject;

            if (val != null)
                val.SetNull();
        }

        /// <summary>
        /// Determines whether the given DocumentObject is null (not set).
        /// </summary>
        public override bool IsNull(DocumentObject dom)
        {
            DocumentObject val = InternalGetValue(dom) as DocumentObject;

            return val == null || val.IsNull();
        }
    }

    /// <summary>
    /// Value descriptor of DocumentObjectCollection.
    /// </summary>
    internal class DocumentObjectCollectionDescriptor : ValueDescriptor
    {
        internal DocumentObjectCollectionDescriptor(string valueName, Type valueType, Type fieldType, MemberInfo memberInfo, VDFlags flags)
            : base(valueName, valueType, fieldType, memberInfo, flags)
        { }

        public override object GetValue(DocumentObject dom, GV flags)
        {
            Debug.Assert(Enum.IsDefined(typeof(GV), flags), DomSR.InvalidEnumValue(flags));

            Debug.Assert(MemberInfo is FieldInfo, "Properties of DocumentObjectCollection not allowed.");
            DocumentObjectCollection val = InternalGetValue(dom) as DocumentObjectCollection;
            if (val == null && flags == GV.ReadWrite)
            {
                val = (DocumentObjectCollection)CreateValue();
                val._parent = dom;
                InternalSetValue(dom, val);
                return val;
            }
            if (val != null && val.IsNull() && flags == GV.GetNull)
                return null;
            return val;
        }

        public override void SetValue(DocumentObject dom, object val)
        {
            InternalSetValue(dom, val);
        }

        public override void SetNull(DocumentObject dom)
        {
            DocumentObjectCollection val = InternalGetValue(dom) as DocumentObjectCollection;
            if (val != null)
                val.SetNull();
        }

        /// <summary>
        /// Determines whether the given DocumentObject is null (not set).
        /// </summary>
        public override bool IsNull(DocumentObject dom)
        {
            DocumentObjectCollection val = InternalGetValue(dom) as DocumentObjectCollection;
            if (val == null)
                return true;
            return val.IsNull();
        }
    }
}
