package com.chenbitao.word.util;

import javax.imageio.ImageIO;
import java.awt.image.BufferedImage;
import java.io.ByteArrayOutputStream;
import java.io.IOException;

/**
 * Image byte generation helpers for tests and playground demos.
 *
 * <p>Centralizes lightweight {@link BufferedImage} and {@link ImageIO} boilerplate so sample
 * image generation stays consistent across modules.</p>
 */
public final class ImageBytesUtils {

    private ImageBytesUtils() {
    }

    /**
     * Creates a solid-color PNG image.
     *
     * @param width image width in pixels
     * @param height image height in pixels
     * @param rgb RGB color value
     * @return PNG image bytes
     * @throws IOException if the image cannot be written
     */
    public static byte[] solidPng(int width, int height, int rgb) throws IOException {
        return solidImage("png", width, height, rgb);
    }

    /**
     * Creates a solid-color image in the requested format.
     *
     * @param formatName image format, for example {@code png}, {@code jpg}, or {@code bmp}
     * @param width image width in pixels
     * @param height image height in pixels
     * @param rgb RGB color value
     * @return image bytes
     * @throws IOException if the image cannot be written
     */
    public static byte[] solidImage(String formatName, int width, int height, int rgb) throws IOException {
        return image(formatName, width, height, (x, y, imageWidth, imageHeight) -> rgb);
    }

    /**
     * Creates a PNG image using a pixel-level color provider.
     *
     * @param width image width in pixels
     * @param height image height in pixels
     * @param colorProvider provider for each pixel color
     * @return PNG image bytes
     * @throws IOException if the image cannot be written
     */
    public static byte[] png(int width, int height, PixelColorProvider colorProvider) throws IOException {
        return image("png", width, height, colorProvider);
    }

    /**
     * Creates an image in the requested format using a pixel-level color provider.
     *
     * @param formatName image format, for example {@code png}, {@code jpg}, or {@code bmp}
     * @param width image width in pixels
     * @param height image height in pixels
     * @param colorProvider provider for each pixel color
     * @return image bytes
     * @throws IOException if the image cannot be written
     */
    public static byte[] image(String formatName, int width, int height, PixelColorProvider colorProvider) throws IOException {
        BufferedImage image = new BufferedImage(width, height, BufferedImage.TYPE_INT_RGB);
        for (int x = 0; x < image.getWidth(); x++) {
            for (int y = 0; y < image.getHeight(); y++) {
                image.setRGB(x, y, colorProvider.rgb(x, y, image.getWidth(), image.getHeight()));
            }
        }

        ByteArrayOutputStream output = new ByteArrayOutputStream();
        if (!ImageIO.write(image, formatName, output)) {
            throw new IOException("Unsupported image format: " + formatName);
        }
        return output.toByteArray();
    }

    /**
     * Provides RGB values for generated image pixels.
     */
    public interface PixelColorProvider {

        /**
         * Returns the RGB value for the requested pixel.
         *
         * @param x horizontal coordinate
         * @param y vertical coordinate
         * @param width image width in pixels
         * @param height image height in pixels
         * @return RGB color value
         */
        int rgb(int x, int y, int width, int height);
    }
}
