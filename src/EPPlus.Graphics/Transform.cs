/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  27/11/2025         EPPlus Software AB           EPPlus 9
 *************************************************************************************************/
using System;
using System.Collections.Generic;
using System.Linq;
using System.Globalization;
using EPPlus.Graphics.Math;
using OfficeOpenXml;

namespace EPPlus.Graphics
{
    internal class Transform
    {
        public string Name;
        public ExcelRangeBase cell;

        public Vector2 LocalPosition { get; set; } = Vector2.Zero;

        public Vector2 Position
        {
            get
            {
                return TransformPointToWorld(Vector2.Zero);
            }
            set
            {
                if (Parent == null)
                {
                    LocalPosition = value;
                }
                else
                {
                    LocalPosition = value * Parent.GetWorldMatrix().Inverse();
                }
            }
        }

        public Vector2 LocalScale { get; set; } = Vector2.One;

        public Vector2 Scale
        {
            get
            {
                if (Parent == null)
                    return LocalScale;
                else
                    return Parent.Scale * LocalScale;
            }
            set
            {
                if (Parent == null)
                {
                    LocalScale = value;
                }
                else
                {
                    LocalScale = value / Parent.Scale;
                }
            }
        }

        public double LocalRotationRadians => LocalRotation * System.Math.PI / 180.0;

        private double localRotationDegrees = 0;

        public double LocalRotation
        {
            get
            {
                return localRotationDegrees;
            }
            set
            {
                localRotationDegrees = value;
            }
        }

        public double RotationRadians => Rotation * System.Math.PI / 180.0;

        public double Rotation
        {
            get
            {
                return Parent != null ? Parent.Rotation + LocalRotation : LocalRotation;
            }
            set
            {
                if (Parent != null)
                    LocalRotation = value - Parent.Rotation;
                else
                    LocalRotation = value;
            }
        }

        public Vector2 Size { get; set; } = Vector2.Zero;

        public int Z { get; set; } = 0;

        private Transform _parent = null;
        public Transform Parent
        {
            get
            {
                return _parent;
            }
            set
            {
                if (_parent == value) return;
                if (_parent != null)
                {
                    _parent.RemoveChild(this);
                }
                if (value != null)
                {
                    value.AddChild(this);
                }
                else
                {
                    _parent = null;
                }
            }
        }

        private List<Transform> _childObjects = null;
        public List<Transform> ChildObjects
        {
            get
            {
                if (_childObjects == null)
                {
                    _childObjects = new List<Transform>();
                }
                return _childObjects;
            }
            set
            {
                if (_childObjects == null)
                {
                    _childObjects = new List<Transform>();
                }
            }
        }

        public Transform() { }

        public Transform(Vector2 position, Vector2 size)
            : this(position, size, Vector2.One, 0d, null) { }

        public Transform(Vector2 position, Vector2 size, Vector2 scale)
            : this(position, size, scale, 0d, null) { }

        public Transform(Vector2 position, Vector2 size, Vector2 scale, double rotation)
            : this(position, size, scale, rotation, null) { }

        public Transform(Vector2 position, Vector2 size, Transform parent)
            : this(position, size, Vector2.One, 0d, null) { }

        public Transform(Vector2 position, Vector2 size, Vector2 scale, Transform parent)
            : this(position, size, scale, 0d, null) { }

        public Transform(Vector2 position, Vector2 size, Vector2 scale, double rotation, Transform parent)
        {
            Position = position;
            Size = size;
            Scale = scale;
            Rotation = rotation;
            Parent = parent;
            ChildObjects = null;
        }

        public Transform(double x, double y, double width, double height, double scaleX = 1, double scaleY = 1, double rotation = 0, Transform parent = null)
        {
            Position = new Vector2(x, y);
            Size = new Vector2(width, height);
            Scale = new Vector2(scaleX, scaleY);
            Rotation = rotation;
            Parent = parent;
            ChildObjects = null;
        }
        public Transform AddChild(Transform child)
        {
            Vector2 worldPos;
            if(child.Parent != null)
            {
                worldPos = child.Position;
                child.Parent.RemoveChild(child);
                var parentInverse = GetWorldMatrix().Inverse();
                child.LocalPosition = worldPos * parentInverse;
            }
            if (!ChildObjects.Contains(child))
            {
                ChildObjects.Add(child);
            }
            child._parent = this;
            return child;
        }

        public void RemoveChild(Transform child)
        {
            if(ChildObjects.Remove(child))
            {
                child._parent = null;
            }
        }

        public void Translate(Vector2 offset)
        {
            LocalPosition += offset;
        }
        public void Translate(double x, double y)
        {
            Translate(new Vector2(x, y));
        }

        public Vector2 TransformPointToLocal(Vector2 point)
        {
            return point * GetWorldMatrix().Inverse();
        }

        public Vector2 TransformPointToWorld(Vector2 point)
        {
            return point * GetWorldMatrix();
        }

        public Matrix3x3 GetLocalMatrix()
        {
            var scale = Matrix3x3.Scaling(LocalScale.X, LocalScale.Y);
            var rotation = Matrix3x3.Rotation(LocalRotation);
            var translation = Matrix3x3.Translation(LocalPosition.X, LocalPosition.Y);
            return scale * rotation * translation;
        }

        public Matrix3x3 GetWorldMatrix()
        {
            return Parent != null ? GetLocalMatrix() * Parent.GetWorldMatrix() : GetLocalMatrix();
        }

        public Rect GetGlobalBoundingbox()
        {
            var worldMatrix = GetWorldMatrix();
            var corners = new[] { new Vector2(0, 0), new Vector2(Size.X, 0), new Vector2(0, Size.Y), new Vector2(Size.X, Size.Y) }.Select(p => p * worldMatrix);
            var minX = corners.Min(p => p.X);
            var minY = corners.Min(p => p.Y);
            var maxX = corners.Max(p => p.X);
            var maxY = corners.Max(p => p.Y);
            var rect = new Rect();
            rect.X = minX;
            rect.Y = minY;
            rect.Width = maxX - minX;
            rect.Height = maxY - minY;
            rect.Top = rect.Y;
            rect.Left = rect.X;
            rect.Bottom = rect.Y + rect.Height;
            rect.Right = rect.X + rect.Width;
            return rect;
        }

        public static bool Intersects(Rect bbox, Rect pageBounds)
        {
            return !(bbox.Right  < pageBounds.Left    ||
                     bbox.Left   > pageBounds.Right   ||
                     bbox.Bottom < pageBounds.Top     ||
                     bbox.Top    > pageBounds.Bottom  );
        }

        public static bool IntersectsFully(Rect contentBounds, Rect cellBounds)
        {
            return cellBounds.Left   >= contentBounds.Left  &&
                   cellBounds.Top    >= contentBounds.Top   &&
                   cellBounds.Right  <= contentBounds.Right &&
                   cellBounds.Bottom <= contentBounds.Bottom;
        }

        public string ToHierarchyString(int indentLevel = 0)
        {
            var indent = new string(' ', indentLevel * 4);
            var result = $"{indent}{Name ?? GetType().Name}|" +
                         $"({LocalPosition.X.ToString(CultureInfo.InvariantCulture)},{LocalPosition.Y.ToString(CultureInfo.InvariantCulture)}):" +
                         $"({Size.X.ToString(CultureInfo.InvariantCulture)},{Size.Y.ToString(CultureInfo.InvariantCulture)})";
            foreach (var child in ChildObjects)
            {
                result += Environment.NewLine + child.ToHierarchyString(indentLevel + 1);
            }
            return result;
        }
    }
}
