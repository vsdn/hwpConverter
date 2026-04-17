package kr.n.nframe.test;

import kr.n.nframe.hwplib.writer.DistributionWriter;
import org.apache.poi.poifs.filesystem.*;
import javax.crypto.Cipher;
import javax.crypto.spec.SecretKeySpec;
import java.io.*;
import java.security.MessageDigest;
import java.nio.charset.StandardCharsets;

/**
 * 테스트: 알려진 비밀번호로 reference 배포용 HWP의 ViewText를 복호화한다.
 */
public class DistDecryptTest {
    public static void main(String[] args) throws Exception {
        String distPath = args[0];
        String password = args[1];

        POIFSFileSystem fs = new POIFSFileSystem(new FileInputStream(distPath));
        DirectoryEntry vt = (DirectoryEntry) fs.getRoot().getEntry("ViewText");
        byte[] viewData = StreamExtractor.readStream(vt, "Section0");

        System.out.println("ViewText/Section0: " + viewData.length + " bytes");

        // DISTRIBUTE_DOC_DATA 레코드 헤더 파싱
        int header = le32(viewData, 0);
        int tagId = header & 0x3FF;
        int size = (header >> 20) & 0xFFF;
        System.out.printf("Record: tag=0x%03X size=%d%n", tagId, size);

        // 256바이트 페이로드 추출 (4바이트 헤더 뒤)
        byte[] distData = new byte[256];
        System.arraycopy(viewData, 4, distData, 0, 256);

        // seed 추출 (페이로드의 처음 4바이트)
        int seed = le32(distData, 0);
        System.out.printf("Seed: 0x%08X%n", seed);

        // 랜덤 배열 생성
        byte[] randomArray = DistributionWriter.buildRandomArray(seed);

        // XOR로 디코드
        byte[] decoded = new byte[256];
        for (int i = 0; i < 256; i++) decoded[i] = (byte)(distData[i] ^ randomArray[i]);

        // offset 위치의 해시 추출
        int offset = (seed & 0x0F) + 4;
        System.out.println("Offset: " + offset);

        byte[] storedHash = new byte[80];
        System.arraycopy(decoded, offset, storedHash, 0, Math.min(80, 256 - offset));
        System.out.print("Stored hash (first 20 bytes): ");
        for (int i = 0; i < 20; i++) System.out.printf("%02X ", storedHash[i] & 0xFF);
        System.out.println();

        int flags = (decoded[offset + 80] & 0xFF) | ((decoded[offset + 81] & 0xFF) << 8);
        System.out.printf("Flags: 0x%04X (noCopy=%b, noPrint=%b)%n", flags, (flags&1)!=0, (flags&2)!=0);

        // 비밀번호의 SHA1 계산
        byte[] pwSha1 = MessageDigest.getInstance("SHA-1").digest(
                password.getBytes(StandardCharsets.UTF_16LE));
        System.out.print("Password SHA1: ");
        for (int i = 0; i < 20; i++) System.out.printf("%02X ", pwSha1[i] & 0xFF);
        System.out.println();

        // 비교
        boolean match = true;
        for (int i = 0; i < 20; i++) {
            if (storedHash[i] != pwSha1[i]) { match = false; break; }
        }
        System.out.println("SHA1 match: " + match);

        // 비밀번호 SHA1 키로 복호화 시도
        byte[] aesKey = new byte[16];
        System.arraycopy(pwSha1, 0, aesKey, 0, 16);

        byte[] encrypted = new byte[viewData.length - 260];
        System.arraycopy(viewData, 260, encrypted, 0, encrypted.length);
        System.out.println("Encrypted data: " + encrypted.length + " bytes");

        Cipher cipher = Cipher.getInstance("AES/ECB/NoPadding");
        cipher.init(Cipher.DECRYPT_MODE, new SecretKeySpec(aesKey, "AES"));
        byte[] decrypted = cipher.doFinal(encrypted);

        System.out.println("Decrypted: " + decrypted.length + " bytes");
        System.out.print("First 32 bytes: ");
        for (int i = 0; i < 32; i++) System.out.printf("%02X ", decrypted[i] & 0xFF);
        System.out.println();

        // zlib 압축 해제 시도
        try {
            byte[] decompressed = StreamExtractor.decompress(decrypted);
            System.out.println("Decompressed: " + decompressed.length + " bytes (SUCCESS)");
        } catch (Exception e) {
            System.out.println("Decompression failed: " + e.getMessage());
            // 압축되지 않았을 수 있음, 레코드로 파싱 시도
            int rh = le32(decrypted, 0);
            System.out.printf("First record: tag=0x%03X level=%d size=%d%n",
                    rh & 0x3FF, (rh>>10)&0x3FF, (rh>>20)&0xFFF);
        }

        fs.close();
    }

    static int le32(byte[] d, int o) {
        return (d[o]&0xFF)|((d[o+1]&0xFF)<<8)|((d[o+2]&0xFF)<<16)|((d[o+3]&0xFF)<<24);
    }
}
