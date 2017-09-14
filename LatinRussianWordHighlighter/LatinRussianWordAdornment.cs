using System;
using System.Text.RegularExpressions;
using System.Windows.Controls;
using System.Windows.Media;
using Microsoft.VisualStudio.Text;
using Microsoft.VisualStudio.Text.Editor;
using Microsoft.VisualStudio.Text.Formatting;

namespace VSIXProject1
{
    /// <summary>
    /// Highlight words with mixed latin and rus characters
    /// </summary>
    internal sealed class LatinRussianWordHightliterAdornment
    {
        /// <summary>
        /// The layer of the adornment.
        /// </summary>
        private readonly IAdornmentLayer layer;

        /// <summary>
        /// Text view where the adornment is created.
        /// </summary>
        private readonly IWpfTextView view;

        /// <summary>
        /// Adornment brush.
        /// </summary>
        private readonly Brush brush;

        /// <summary>
        /// Adornment pen.
        /// </summary>
        private readonly Pen pen;

        /// <summary>
        /// Initializes a new instance of the <see cref="LatinRussianWordHightliterAdornment"/> class.
        /// </summary>
        /// <param name="view">Text view to create the adornment for</param>
        public LatinRussianWordHightliterAdornment(IWpfTextView view)
        {
            if (view == null)
            {
                throw new ArgumentNullException("view");
            }

            this.layer = view.GetAdornmentLayer("TextAdornment1");

            this.view = view;
            this.view.LayoutChanged += this.OnLayoutChanged;

            // Create the pen and brush to color the box behind the a's
            this.brush = new SolidColorBrush(Color.FromArgb(0x20, 0x00, 0x00, 0xff));
            this.brush.Freeze();

            var penBrush = new SolidColorBrush(Colors.DarkBlue);
            penBrush.Freeze();
            this.pen = new Pen(penBrush, 0.5);
            this.pen.Freeze();
        }

        /// <summary>
        /// Handles whenever the text displayed in the view changes by adding the adornment to any reformatted lines
        /// </summary>
        /// <remarks><para>This event is raised whenever the rendered text displayed in the <see cref="ITextView"/> changes.</para>
        /// <para>It is raised whenever the view does a layout (which happens when DisplayTextLineContainingBufferPosition is called or in response to text or classification changes).</para>
        /// <para>It is also raised whenever the view scrolls horizontally or when its size changes.</para>
        /// </remarks>
        /// <param name="sender">The event sender.</param>
        /// <param name="e">The event arguments.</param>
        internal void OnLayoutChanged(object sender, TextViewLayoutChangedEventArgs e)
        {
            foreach (var line in e.NewOrReformattedLines)
            {
                this.CreateVisuals(line);
            }

            //this.CreateVisuals();
        }

        private void CreateVisuals(ITextViewLine line)
        {
            IWpfTextViewLineCollection textViewLines = view.TextViewLines;
            TryGetText(view, line, out string text);

            var regex = new Regex(@"\b(?=[а-яА-ЯёЁ]*[a-zA-Z])(?=[a-zA-Z]*[а-яА-ЯёЁ])[\wа-яА-ЯёЁ]+\b");
            var match = regex.Match(text);

            while (match.Success)
            {
                var matchStart = line.Start.Position + match.Index;
                var span = new SnapshotSpan(view.TextSnapshot, Span.FromBounds(matchStart, matchStart + match.Length));
                DrawAdornment(textViewLines, span);
                match = match.NextMatch();
            }
        }


        /// <summary>
        /// This will get the text of the ITextView line as it appears in the actual user editable 
        /// document. 
        /// jared parson: https://gist.github.com/4320643
        /// </summary>
        public static bool TryGetText(IWpfTextView textView, ITextViewLine textViewLine, out string text)
        {
            var extent = textViewLine.Extent;
            var bufferGraph = textView.BufferGraph;
            try
            {
                var collection = bufferGraph.MapDownToSnapshot(extent, SpanTrackingMode.EdgeInclusive, textView.TextSnapshot);
                var span = new SnapshotSpan(collection[0].Start, collection[collection.Count - 1].End);
                text = span.GetText();
                return true;
            }
            catch
            {
                text = null;
                return false;
            }
        }

        private void DrawAdornment(IWpfTextViewLineCollection textViewLines, SnapshotSpan span)
        {
            Geometry geometry = textViewLines.GetMarkerGeometry(span);
            if (geometry != null)
            {
                var drawing = new GeometryDrawing(this.brush, this.pen, geometry);
                drawing.Freeze();

                var drawingImage = new DrawingImage(drawing);
                drawingImage.Freeze();

                var image = new Image
                {
                    Source = drawingImage,
                };

                // Align the image with the top of the bounds of the text geometry
                Canvas.SetLeft(image, geometry.Bounds.Left);
                Canvas.SetTop(image, geometry.Bounds.Top);

                this.layer.AddAdornment(AdornmentPositioningBehavior.TextRelative, span, null, image, null);
            }
        }
    }
}
