using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using AForge;
using AForge.Math.Geometry;

namespace LitePlacer
{
	public class Shapes
	{
		public class Component
		{
			public AForge.Point Center { get; set; }	// Centerpoint of component
			public decimal Alignment { get; set; }		// angle of component
			public List<IntPoint> Outline { get; set; }
			public LineSegment Longest { get; set; }	// Longest line segment in Outline (needed in drawing, avoid calculating twice)
			public AForge.Point NormalStart { get; set; }  // (needed in drawing, avoid calculating twice)
			public AForge.Point NormalEnd { get; set; }		// (needed in drawing, avoid calculating twice)

			public Component(AForge.Point centr, decimal alignmnt, List<IntPoint> outln, 
							 LineSegment lngst, AForge.Point Nstart, AForge.Point Nend)
			{
				Center = centr;
				Alignment = alignmnt;
				Outline = outln;
				Longest = lngst;
				NormalStart = Nstart;
				NormalEnd = Nend;
			}
		}

		public class Circle
		{
			public decimal X { get; set; }
			public decimal Y { get; set; }
			public decimal Radius { get; set; }

			public Circle(double x, double y, double r)
			{
				X = (decimal)x;
				Y = (decimal)y;
				Radius = (decimal)r;
			}
		}

	}


}
