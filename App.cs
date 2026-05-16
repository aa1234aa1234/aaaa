using Microsoft.VisualBasic;
using ohmygod.EventListeners;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Threading.Tasks;

namespace ohmygod
{
    internal class App
    {
        private TradingEngine engine;
        private StartUpListener StartUpListener;
        private BuyOrderListener BuyOrderListener;
        private SellOfferListener SellOfferListener;

        public App() {
            engine = new();
            StartUpListener = new("STARTUP", engine);
            BuyOrderListener = new("SETBUYORDER", engine);
            SellOfferListener = new("SETSELLOFFER", engine);
        }

        public void run(Form1 form)
        {
            if(engine.GetUuid() == string.Empty)
            {
                loginPrompt();
            }
            engine.run(form);
        }

        public void loginPrompt()
        {
            string info = Interaction.InputBox("account info");
            engine.SetUuid(CreateUuid(info).ToString());
            MessageBox.Show(engine.GetUuid());
        }

        private Guid CreateUuid(string info)
        {
            Guid namespaceGuid = new Guid("12345678-1234-1234-1234-123456789abc");

            byte[] namespaceBytes = namespaceGuid.ToByteArray();
            byte[] inputBytes = Encoding.UTF8.GetBytes(info);

            byte[] combined = new byte[namespaceBytes.Length + inputBytes.Length];

            Buffer.BlockCopy(namespaceBytes, 0, combined, 0, namespaceBytes.Length);
            Buffer.BlockCopy(inputBytes, 0, combined, namespaceBytes.Length, inputBytes.Length);

            using SHA1 sha1 = SHA1.Create();

            byte[] hash = sha1.ComputeHash(combined);

            byte[] newGuid = new byte[16];
            Array.Copy(hash, newGuid, 16);

            newGuid[6] = (byte)((newGuid[6] & 0x0F) | 0x50);
            newGuid[8] = (byte)((newGuid[8] & 0x3F) | 0x80);

            return new Guid(newGuid);
        }
    }
}
