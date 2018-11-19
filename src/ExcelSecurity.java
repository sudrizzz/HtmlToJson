import org.bouncycastle.crypto.BufferedBlockCipher;
import org.bouncycastle.crypto.engines.AESFastEngine;
import org.bouncycastle.crypto.modes.CBCBlockCipher;
import org.bouncycastle.crypto.paddings.PaddedBufferedBlockCipher;
import org.bouncycastle.crypto.params.KeyParameter;
import org.bouncycastle.crypto.params.ParametersWithIV;
import org.bouncycastle.util.encoders.Hex;

public class ExcelSecurity {

    static byte[] keybytes = "WEAVER E-DESIGN.".getBytes();
    static byte[] iv = "weaver e-design.".getBytes();

    /**
     *为字符串加密
     *
     * @param content
     * @return
     */
    public String encodeStr(String content) {
        try {
            byte[] keybytes = "WEAVER E-DESIGN.".getBytes();
            byte[] iv = "weaver e-design.".getBytes();
            BufferedBlockCipher engine = new PaddedBufferedBlockCipher(
                    new CBCBlockCipher(new AESFastEngine()));
            engine.init(true, new ParametersWithIV(new KeyParameter(keybytes),
                    iv));
            byte[] enc = new byte[engine
                    .getOutputSize(content.getBytes().length)];
            int size1 = engine.processBytes(content.getBytes(), 0, content
                    .getBytes().length, enc, 0);
            int size2 = engine.doFinal(enc, size1);
            byte[] encryptedContent = new byte[size1 + size2];
            System.arraycopy(enc, 0, encryptedContent, 0,
                    encryptedContent.length);
            return new String(Hex.encode(encryptedContent));
        } catch (Exception ex) {
            ex.printStackTrace();
        }
        return "";
    }

    /**
     * 为字符串解密
     *
     * @param content
     * @return
     */
    public String decode(String content) {
        try {
            BufferedBlockCipher engine = new PaddedBufferedBlockCipher(
                    new CBCBlockCipher(new AESFastEngine()));
            engine.init(true, new ParametersWithIV(new KeyParameter(keybytes),
                    iv));
            byte[] deByte = Hex.decode(content);
            engine.init(false, new ParametersWithIV(new KeyParameter(keybytes),
                    iv));
            byte[] dec = new byte[engine.getOutputSize(deByte.length)];
            int size1 = engine.processBytes(deByte, 0, deByte.length, dec, 0);
            int size2 = engine.doFinal(dec, size1);
            byte[] decryptedContent = new byte[size1 + size2];
            System.arraycopy(dec, 0, decryptedContent, 0,
                    decryptedContent.length);
            return new String(decryptedContent);
        } catch (Exception ex) {
            ex.printStackTrace();
        }
        return "";
    }

}
