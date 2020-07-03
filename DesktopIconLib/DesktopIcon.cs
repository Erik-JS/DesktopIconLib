using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DesktopIconLib
{
    public class DesktopIcon
    {
        private int x;
        private int y;
        private string label;

        public DesktopIcon(string IconText, int IconX, int IconY)
        {
            x = IconX;
            y = IconY;
            label = IconText;
        }

        public Tuple<int, int> GetCoordinates()
        {
            return new Tuple<int, int>(x, y);
        }

        public void GetCoordinates(out int X, out int Y)
        {
            X = x;
            Y = y;
        }

        public void SetCoordinates(int X, int Y)
        {
            DesktopInfo.SetIconCoordinates(label, X, Y);
            x = X;
            y = Y;
        }

        public void RestorePosition()
        {
            DesktopInfo.SetIconCoordinates(label, x, y);
        }

        public string Label
        {
            get { return label; }
        }

        public int X
        {
            get { return x; }
        }

        public int Y
        {
            get { return y; }
        }

    }
}
