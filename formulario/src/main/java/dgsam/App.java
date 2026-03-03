package dgsam;

import javax.swing.SwingUtilities;

public class App {
    public static void main(String[] args) {
        try {
            com.formdev.flatlaf.FlatLightLaf.setup();
        } catch (Exception ignored) {}

        SwingUtilities.invokeLater(() -> {
            ListProvider provider = ListProvider.demo();
            FichaDatosPersonalesFrame frame = new FichaDatosPersonalesFrame(provider);
            frame.setVisible(true);
        });
    }
}
