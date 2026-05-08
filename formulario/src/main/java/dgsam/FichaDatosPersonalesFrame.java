package dgsam;
//import com.google.gson.Gson;
//import com.google.gson.JsonObject;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import javax.imageio.ImageIO;
import javax.swing.*;
import javax.swing.border.EmptyBorder;
import javax.swing.plaf.basic.BasicTabbedPaneUI;
import javax.swing.text.AbstractDocument;
import javax.swing.text.AttributeSet;
import javax.swing.text.BadLocationException;
import javax.swing.text.DocumentFilter;
/*import javax.swing.text.DateFormatter;
import javax.swing.text.DefaultFormatterFactory;*/
import java.awt.BasicStroke;
import java.awt.BorderLayout;
import java.awt.Color;
import java.awt.Dimension;
import java.awt.Rectangle;
import java.awt.FlowLayout;
import java.awt.Font;
import java.awt.Graphics;
import java.awt.Graphics2D;
import java.awt.Insets;
import java.awt.LayoutManager;
import java.awt.RenderingHints;
import java.awt.event.ActionEvent;
import java.awt.image.BufferedImage;
import java.awt.GridBagConstraints;
import java.awt.GridBagLayout;
import java.awt.Image;
import java.io.BufferedReader;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.InputStreamReader;
/*import java.net.HttpURLConnection;
import java.net.URL;*/
import java.nio.charset.StandardCharsets;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Arrays;
//import java.util.Calendar;
import java.util.Collections;
import java.util.Date;
import java.util.HashMap;
import java.util.List;
import java.util.Locale;
import java.util.Map;
import java.util.Objects;
import java.util.Set;
import java.util.TreeSet;
import java.util.function.Supplier;

import com.toedter.calendar.JDateChooser;
import com.toedter.calendar.JTextFieldDateEditor;

public class FichaDatosPersonalesFrame extends JFrame {

    private static final Color FORM_ROOT_BG = new Color(0xE8F2FF);
    private static final Color FORM_CARD_BORDER = new Color(0xC8DCFF);
    private static final Color FORM_HEADER_BG = new Color(0x0D6EFD);
    private static final Color FORM_HEADER_TEXT = Color.WHITE;
    private static final Color FORM_LABEL_TEXT = new Color(0x1D3A6A);
    private static final Color FORM_BUTTON_PRIMARY = new Color(0x0D6EFD);
    private static final Color FORM_BUTTON_SECONDARY = new Color(0x6C757D);

    private final ListProvider listProvider;
    private final FichaModel model = new FichaModel();

    // Mapeos para cascada (se cargan desde Excel si existe)
    private final Map<String, Set<String>> escalafon_especBasicaMap = new HashMap<>();
    private final Map<String, Set<String>> especBasica_especAvanzadaMap = new HashMap<>();

    // ===== Campos básicos =====
    private JTextField txtIOSFA;
    private JTextField txtDNI;
    private JComboBox<String> cmbGrado;
    private JTextField txtApellido;
    private JTextField txtNombres;
    private JTextField txtLugarNac;
    private JTextField txtCUIL;
    private JTextField txtCBU;
    private JTextField txtCelular;
    private JTextField txtCP;
    private JComboBox<String> cmbProvincia;

    // Datos personales (FECHAS: PICKER)
    private JDateChooser spFechaIngreso;
    private JTextField txtAniosEnGrado;
    private JTextField txtAniosEnEspecialidad;
    private JComboBox<String> cmbSeDesempena;
    private JComboBox<String> cmbCumpleTurno;
    private JTextField txtFuncion;
    private JTextField txtPromedioTurnos;
    private JDateChooser spAptitudPsicofisicaFecha;
    private JDateChooser spFechaCondicionTiro;
    private JComboBox<String> cmbPoseeAptoFisico;

    private JComboBox<String> cmbEspBasica;
    private JComboBox<String> cmbEspAvanzada;
    private JComboBox<String> cmbEscalafon;

    private JTextField txtEmailInst;
    private JDateChooser spFechaNacimiento;
    private JTextField txtUsuarioGDE;
    private JTextField txtRTI;
    private JComboBox<String> cmbFactorSanguineo;
    private JComboBox<String> cmbUnidadRevista;
    private JTextField txtDestinoInterno;
    private JTextField txtCargo;
    private JComboBox<String> cmbDestinoAnterior;
    private JComboBox<String> cmbDeseaPermanecer;

    // Idiomas
    private JTextField txtIdioma1;
    private JTextField txtNivelIdioma1;
    private JDateChooser spFechaNivel1;
    private JTextField txtIdioma2;
    private JTextField txtNivelIdioma2;
    private JDateChooser spFechaNivel2;
    private JTextField txtIdioma3;
    private JTextField txtNivelIdioma3;
    private JDateChooser spFechaNivel3;
    private JComboBox<String> cmbRindioSidiel;

    // Capacitaciones y docencia
    private JTextField txtCapInstTitulo;
    private JTextField txtCapInstExpedidoPor;
    private JDateChooser spCapInstFecha;
    private JTextField txtMaxCapExtraTitulo;
    private JTextField txtMaxCapExtraExpedidoPor;
    private JDateChooser spMaxCapExtraFecha;
    private JComboBox<String> cmbActividadProfesor;
    private JTextField txtAsignaturaTemas;
    private JComboBox<String> cmbDictaActualmente;
    private JTextField txtModalidad;
    private JTextField txtTituloHabilitante;

    // Comisiones / campañas
    private JTextField txtComisionExteriorMotivo;
    private JTextField txtComisionExteriorPaisCiudad;
    private JDateChooser spFechaInicioComision;
    private JDateChooser spFechaFinComision;
    private JComboBox<String> cmbCumplioCampanasAntarticas;
    private JTextField txtCantidadCampanas;
    private JTextField txtDotacionGpoTareas;
    private JTextField txtCargoDesempenado;
    private JDateChooser spPeriodoDesde;
    private JDateChooser spPeriodoHasta;

    // ===== Domicilio =====
    private JTextField txtDomicilioCalle;
    private JTextField txtNumeroCalle;
    private JTextField txtLocalidad;

    // ===== Preferencias =====
    private JComboBox<String> cmbEstadoCivil;
    private JComboBox<String> cmbConyugeGrado;        // (lo mantenemos por compatibilidad si lo usás en otros lados)
    private JComboBox<String> cmbDestino1;
    private JComboBox<String> cmbDestino2;
    private JTextField txtObservEstadoCivil;

    // Cónyuge (campos base)
    private JTextField txtConyugeApellido;
    private JTextField txtConyugeNombre;
    private JDateChooser spConyugeFechaNacimiento;
    private JTextField txtConyugeDNI;

    // Cónyuge militar? + campos condicionados
    private JComboBox<String> cmbConyugeEsMilitar;    // Sí/No
    private JTextField txtConyugeNroId;               // solo si militar
   private JComboBox<String> cmbConyugeDestino;      // solo si militar
    // Cónyuge especialidad (cascada) solo si militar
    private JComboBox<String> cmbConyugeEscalafon;
    private JComboBox<String> cmbConyugeEspBasica;
    private JComboBox<String> cmbConyugeEspAvanzada;

    // Hijos
    private JComboBox<String> cmbHijos;
    private JTextField txtCantidadHijos;

    // TEXT AREA
    private JTextArea taImpedimentoTraslado;

    // ====== LOGO (mejorado) ======
    private static final String LOGO_RESOURCE = "/logoDGSAM.png";

    public FichaDatosPersonalesFrame(ListProvider listProvider) {
        super("Formulario de Datos Personales - DGSAM");
        this.listProvider = Objects.requireNonNull(listProvider, "listProvider");

        // Icono de la ventana
        ImageIcon frameIcon = loadLogoIcon(LOGO_RESOURCE, 32, 32);
        if (frameIcon != null) setIconImage(frameIcon.getImage());

        setDefaultCloseOperation(WindowConstants.DISPOSE_ON_CLOSE);
        setMinimumSize(new Dimension(1100, 700));
        setLocationRelativeTo(null);

        setContentPane(buildRoot());
        loadComboData();

        // cascada principal
        cmbEspBasica.setEnabled(false);
        cmbEspAvanzada.setEnabled(false);
        setupCascadeListeners();
    }

    private JComponent buildRoot() {
        JPanel root = new JPanel(new BorderLayout(0, 0));
        root.setBorder(new EmptyBorder(18, 18, 18, 18));
        root.setBackground(FORM_ROOT_BG);

        JPanel card = createRoundedPanel(new BorderLayout(0, 12), Color.WHITE, 18, FORM_CARD_BORDER, 1);
        card.setBorder(new EmptyBorder(16, 16, 16, 16));

        // Header: título + logo prolijo
        JPanel headerPanel = new JPanel(new BorderLayout());
        headerPanel.setBackground(FORM_HEADER_BG);

        JLabel title = new JLabel("Formulario de Datos Personales", SwingConstants.CENTER);
        title.setFont(title.getFont().deriveFont(Font.BOLD, 20f));
        title.setForeground(FORM_HEADER_TEXT);
        title.setHorizontalAlignment(SwingConstants.CENTER);
        headerPanel.add(title, BorderLayout.CENTER);

        JPanel logoWrap = new JPanel(new FlowLayout(FlowLayout.RIGHT, 0, 0));
        logoWrap.setBackground(FORM_HEADER_BG);

        ImageIcon headerLogo = loadLogoIcon(LOGO_RESOURCE, 48, 48);
        if (headerLogo != null) {
            JLabel logoLabel = new JLabel(headerLogo);
            logoLabel.setPreferredSize(new Dimension(48, 48));
            logoLabel.setMinimumSize(new Dimension(48, 48));
            logoLabel.setHorizontalAlignment(SwingConstants.RIGHT);
            logoLabel.setVerticalAlignment(SwingConstants.CENTER);
            logoWrap.add(logoLabel);
        }
        headerPanel.add(logoWrap, BorderLayout.EAST);
        card.add(headerPanel, BorderLayout.NORTH);

        JTabbedPane tabs = new JTabbedPane();
        styleTabbedPane(tabs);
        tabs.addTab("Datos personales", wrapScrollable(buildDatosPanel()));
        tabs.addTab("Domicilio", wrapScrollable(buildDomicilioPanel()));
        tabs.addTab("Preferencias", wrapScrollable(buildPreferenciasPanel()));
        card.add(tabs, BorderLayout.CENTER);

        card.add(buildFooterButtons(), BorderLayout.SOUTH);
        root.add(card, BorderLayout.CENTER);
        return root;
    }

    // =========================
    // LOGO helpers (sin deformar)
    // =========================
    private ImageIcon loadLogoIcon(String resourcePath, int maxW, int maxH) {
        try (InputStream is = FichaDatosPersonalesFrame.class.getResourceAsStream(resourcePath)) {
            if (is == null) return null;
            BufferedImage img = ImageIO.read(is);
            if (img == null) return null;
            Image scaled = scaleToFit(img, maxW, maxH);
            return new ImageIcon(scaled);
        } catch (Exception e) {
            return null;
        }
    }

    private static Image scaleToFit(BufferedImage src, int maxW, int maxH) {
        int w = src.getWidth();
        int h = src.getHeight();
        if (w <= 0 || h <= 0) return src;

        double rw = (double) maxW / (double) w;
        double rh = (double) maxH / (double) h;
        double r = Math.min(rw, rh);

        if (r >= 1.0) return src;

        int nw = Math.max(1, (int) Math.round(w * r));
        int nh = Math.max(1, (int) Math.round(h * r));
        return src.getScaledInstance(nw, nh, Image.SCALE_SMOOTH);
    }

    private JComponent wrapScrollable(JComponent content) {
        JPanel wrapper = new JPanel(new BorderLayout());
        wrapper.setBackground(Color.WHITE);
        wrapper.add(content, BorderLayout.NORTH);

        JScrollPane sp = new JScrollPane(wrapper);
        sp.setBorder(BorderFactory.createEmptyBorder());
        sp.getVerticalScrollBar().setUnitIncrement(18);
        sp.getViewport().setBackground(Color.WHITE);
        return sp;
    }

    // =========================
    // Datos personales (igual)
    // =========================
    private JComponent buildDatosPanel() {
        JPanel panel = new JPanel(new GridBagLayout());
        panel.setBackground(Color.WHITE);
        panel.setBorder(new EmptyBorder(8, 8, 8, 8));

        JPanel top3Cols = new JPanel(new GridBagLayout());
        top3Cols.setBackground(Color.WHITE);

        JPanel colL = new JPanel(new GridBagLayout());
        colL.setBackground(Color.WHITE);

        JPanel colC = new JPanel(new GridBagLayout());
        colC.setBackground(Color.WHITE);

        JPanel colRTop = new JPanel(new GridBagLayout());
        colRTop.setBackground(Color.WHITE);

        JPanel bottomFull = new JPanel(new GridBagLayout());
        bottomFull.setBackground(Color.WHITE);

        // ===== Inicializar controles =====
        txtIOSFA = new JTextField(18);
        txtDNI = new JTextField(18);
        cmbGrado = newCombo();

        txtApellido = new JTextField(18);
        txtNombres = new JTextField(18);
        txtLugarNac = new JTextField(18);

        spFechaIngreso = newDatePicker();
        txtAniosEnGrado = newNumericField(6);
        txtAniosEnEspecialidad = newNumericField(6);
        cmbSeDesempena = newCombo();
        cmbCumpleTurno = newCombo();
        txtFuncion = new JTextField(18);
        txtPromedioTurnos = new JTextField(6);
        spAptitudPsicofisicaFecha = newDatePicker();
        spFechaCondicionTiro = newDatePicker();
        cmbPoseeAptoFisico = newCombo();
        spFechaNacimiento = newDatePicker();

        cmbEspBasica = newCombo();
        cmbEspAvanzada = newCombo();
        cmbEscalafon = newCombo();

        txtCUIL = new JTextField(18);
        txtCBU = new JTextField(18);
        txtEmailInst = new JTextField(18);
        txtCelular = new JTextField(18);
        txtUsuarioGDE = new JTextField(18);
        txtRTI = new JTextField(12);
        cmbFactorSanguineo = newCombo();
        cmbUnidadRevista = newCombo();
        txtDestinoInterno = new JTextField(18);
        txtCargo = new JTextField(18);
        cmbDestinoAnterior = newCombo();
        cmbDeseaPermanecer = newCombo();

        // Idiomas
        txtIdioma1 = new JTextField(12);
        txtNivelIdioma1 = new JTextField(8);
        spFechaNivel1 = newDatePicker();
        txtIdioma2 = new JTextField(12);
        txtNivelIdioma2 = new JTextField(8);
        spFechaNivel2 = newDatePicker();
        txtIdioma3 = new JTextField(12);
        txtNivelIdioma3 = new JTextField(8);
        spFechaNivel3 = newDatePicker();
        cmbRindioSidiel = newCombo();

        // Capacitaciones / docencia
        txtCapInstTitulo = new JTextField(18);
        txtCapInstExpedidoPor = new JTextField(18);
        spCapInstFecha = newDatePicker();
        txtMaxCapExtraTitulo = new JTextField(18);
        txtMaxCapExtraExpedidoPor = new JTextField(18);
        spMaxCapExtraFecha = newDatePicker();
        cmbActividadProfesor = newCombo();
        txtAsignaturaTemas = new JTextField(18);
        cmbDictaActualmente = newCombo();
        txtModalidad = new JTextField(12);
        txtTituloHabilitante = new JTextField(18);

        // Comisiones / campañas
        txtComisionExteriorMotivo = new JTextField(18);
        txtComisionExteriorPaisCiudad = new JTextField(18);
        spFechaInicioComision = newDatePicker();
        spFechaFinComision = newDatePicker();
        cmbCumplioCampanasAntarticas = newCombo();
        txtCantidadCampanas = new JTextField(6);
        txtDotacionGpoTareas = new JTextField(18);
        txtCargoDesempenado = new JTextField(18);
        spPeriodoDesde = newDatePicker();
        spPeriodoHasta = newDatePicker();

        // ===== TOP (3 columnas) =====
        int rL = 0;
        addField(colL, rL++, "IOSFA", txtIOSFA);
        addField(colL, rL++, "DNI", txtDNI);
        addField(colL, rL++, "Fecha de ingreso", spFechaIngreso);
        addField(colL, rL++, "Años en el grado", txtAniosEnGrado);
        addField(colL, rL++, "Años en la especialidad", txtAniosEnEspecialidad);
        addField(colL, rL++, "¿Se desempeña en la especialidad actualmente?", cmbSeDesempena);
        addField(colL, rL++, "¿Cumple turno o servicio?", cmbCumpleTurno);
        addField(colL, rL++, "Grado", cmbGrado);
        addField(colL, rL++, "Apellido", txtApellido);
        addField(colL, rL++, "Nombres", txtNombres);
        addField(colL, rL++, "Lugar de nacimiento", txtLugarNac);
        addField(colL, rL++, "Fecha de nacimiento", spFechaNacimiento);

        int rC = 0;
        addField(colC, rC++, "Escalafón", cmbEscalafon);
        addField(colC, rC++, "Especialidad básica / primaria", cmbEspBasica);
        addField(colC, rC++, "Especialidad avanzada", cmbEspAvanzada);
        addField(colC, rC++, "Función", txtFuncion);
        addField(colC, rC++, "Promedio anual de turnos realizados", txtPromedioTurnos);
        addField(colC, rC++, "Aptitud psicofísica - fecha último examen", spAptitudPsicofisicaFecha);
        addField(colC, rC++, "Fecha última condición de tiro", spFechaCondicionTiro);
        addField(colC, rC++, "¿Posee apto físico s/mapi 5?", cmbPoseeAptoFisico);
        addField(colC, rC++, "CUIL", txtCUIL);
        addField(colC, rC++, "CBU", txtCBU);
        addField(colC, rC++, "Email (sin @faa.mil.ar)", txtEmailInst);
        addField(colC, rC++, "Celular", txtCelular);

        int rRTop = 0;
        addField(colRTop, rRTop++, "Usuario GDE (sin @faa.mil.ar)", txtUsuarioGDE);
        addField(colRTop, rRTop++, "RTI", txtRTI);
        addField(colRTop, rRTop++, "Factor sanguíneo", cmbFactorSanguineo);
        addField(colRTop, rRTop++, "Unidad de revista", cmbUnidadRevista);
        addField(colRTop, rRTop++, "Destino interno", txtDestinoInterno);
        addField(colRTop, rRTop++, "Cargo", txtCargo);
        addField(colRTop, rRTop++, "Destino anterior", cmbDestinoAnterior);
        addField(colRTop, rRTop++, "Desea permanecer en el destino actual?", cmbDeseaPermanecer);

        GridBagConstraints tc = new GridBagConstraints();
        tc.gridy = 0;
        tc.fill = GridBagConstraints.HORIZONTAL;
        tc.anchor = GridBagConstraints.NORTHWEST;

        tc.gridx = 0;
        tc.weightx = 0.33;
        tc.insets = new Insets(0, 0, 0, 12);
        top3Cols.add(colL, tc);

        tc.gridx = 1;
        tc.weightx = 0.34;
        tc.insets = new Insets(0, 12, 0, 12);
        top3Cols.add(colC, tc);

        tc.gridx = 2;
        tc.weightx = 0.33;
        tc.insets = new Insets(0, 12, 0, 0);
        top3Cols.add(colRTop, tc);

        // ===== BOTTOM (ancho completo) =====
        int rb = 0;

        addSectionTitle(bottomFull, rb++, "Idiomas");
        addField(bottomFull, rb++, "Idioma 1", txtIdioma1);
        addField(bottomFull, rb++, "Nivel idioma 1", txtNivelIdioma1);
        addField(bottomFull, rb++, "Fecha nivel 1", spFechaNivel1);
        addField(bottomFull, rb++, "Idioma 2", txtIdioma2);
        addField(bottomFull, rb++, "Nivel idioma 2", txtNivelIdioma2);
        addField(bottomFull, rb++, "Fecha nivel 2", spFechaNivel2);
        addField(bottomFull, rb++, "Idioma 3", txtIdioma3);
        addField(bottomFull, rb++, "Nivel idioma 3", txtNivelIdioma3);
        addField(bottomFull, rb++, "Fecha nivel 3", spFechaNivel3);
        addField(bottomFull, rb++, "Rindió examen SIDIEL?", cmbRindioSidiel);

        addSectionTitle(bottomFull, rb++, "Capacitaciones / Docencia");
        addField(bottomFull, rb++, "Capacitación Institucional - Título", txtCapInstTitulo);
        addField(bottomFull, rb++, "Capacitación Institucional - Expedido por", txtCapInstExpedidoPor);
        addField(bottomFull, rb++, "Capacitación Institucional - Fecha", spCapInstFecha);
        addField(bottomFull, rb++, "Máxima capacitación extra institucional - Título", txtMaxCapExtraTitulo);
        addField(bottomFull, rb++, "Máxima capacitación extra institucional - Expedido por", txtMaxCapExtraExpedidoPor);
        addField(bottomFull, rb++, "Máxima capacitación extra institucional - Fecha", spMaxCapExtraFecha);
        addField(bottomFull, rb++, "Actividad como profesor/instructor?", cmbActividadProfesor);
        addField(bottomFull, rb++, "Asignatura/temas", txtAsignaturaTemas);
        addField(bottomFull, rb++, "Dicta actualmente?", cmbDictaActualmente);
        addField(bottomFull, rb++, "Modalidad", txtModalidad);
        addField(bottomFull, rb++, "Título habilitante", txtTituloHabilitante);

        addSectionTitle(bottomFull, rb++, "Comisiones / Campañas");
        addField(bottomFull, rb++, "Comisión exterior - Motivo", txtComisionExteriorMotivo);
        addField(bottomFull, rb++, "Comisión exterior - País/Ciudad", txtComisionExteriorPaisCiudad);
        addField(bottomFull, rb++, "Fecha inicio comisión", spFechaInicioComision);
        addField(bottomFull, rb++, "Fecha fin comisión", spFechaFinComision);
        addField(bottomFull, rb++, "Cumplió campañas antárticas?", cmbCumplioCampanasAntarticas);
        addField(bottomFull, rb++, "Cantidad campañas", txtCantidadCampanas);
        addField(bottomFull, rb++, "Dotación / GPO tareas", txtDotacionGpoTareas);
        addField(bottomFull, rb++, "Cargo desempeñado", txtCargoDesempenado);
        addField(bottomFull, rb++, "Periodo desde", spPeriodoDesde);
        addField(bottomFull, rb++, "Periodo hasta", spPeriodoHasta);

        // ===== Insertar TOP + BOTTOM =====
        GridBagConstraints c = new GridBagConstraints();
        c.gridx = 0;
        c.gridy = 0;
        c.weightx = 1.0;
        c.fill = GridBagConstraints.HORIZONTAL;
        c.anchor = GridBagConstraints.NORTHWEST;
        c.insets = new Insets(0, 0, 16, 0);
        panel.add(top3Cols, c);

        c.gridx = 0;
        c.gridy = 1;
        c.weightx = 1.0;
        c.fill = GridBagConstraints.HORIZONTAL;
        c.insets = new Insets(0, 0, 0, 0);
        panel.add(bottomFull, c);

        c.gridx = 0;
        c.gridy = 2;
        c.weighty = 1.0;
        c.fill = GridBagConstraints.BOTH;
        panel.add(Box.createVerticalGlue(), c);

        return panel;
    }

    private void addSectionTitle(JPanel host, int row, String title) {
        GridBagConstraints c = new GridBagConstraints();
        c.gridx = 0;
        c.gridy = row * 2;
        c.weightx = 1.0;
        c.fill = GridBagConstraints.HORIZONTAL;
        c.anchor = GridBagConstraints.WEST;
        c.insets = new Insets(18, 0, 6, 0);

        JLabel lbl = new JLabel(title);
        lbl.setFont(lbl.getFont().deriveFont(Font.BOLD, 13f));
        host.add(lbl, c);

        c.gridy = row * 2 + 1;
        c.insets = new Insets(0, 0, 12, 0);
        host.add(new JSeparator(), c);
    }

    // =========================
    // Domicilio (igual)
    // =========================
    private JComponent buildDomicilioPanel() {
        JPanel panel = new JPanel(new GridBagLayout());
        panel.setBackground(Color.WHITE);
        panel.setBorder(new EmptyBorder(8, 8, 8, 8));

        JPanel colL = new JPanel(new GridBagLayout());
        colL.setBackground(Color.WHITE);
        JPanel colR = new JPanel(new GridBagLayout());
        colR.setBackground(Color.WHITE);

        txtDomicilioCalle = new JTextField(18);
        txtNumeroCalle = new JTextField(18);
        txtLocalidad = new JTextField(18);
        txtCP = new JTextField(18);

        cmbProvincia = newCombo();

        int rL = 0;
        addField(colL, rL++, "Domicilio - Calle", txtDomicilioCalle);
        addField(colL, rL++, "Número", txtNumeroCalle);
        addField(colL, rL++, "Localidad", txtLocalidad);
        addField(colL, rL++, "Código postal", txtCP);

        int rR = 0;
        addField(colR, rR++, "Provincia", cmbProvincia);

        GridBagConstraints c = new GridBagConstraints();
        c.gridx = 0; c.gridy = 0;
        c.weightx = 0.5; c.fill = GridBagConstraints.HORIZONTAL;
        c.anchor = GridBagConstraints.NORTHWEST;
        c.insets = new Insets(0, 0, 0, 12);
        panel.add(colL, c);

        c.gridx = 1; c.gridy = 0;
        c.insets = new Insets(0, 12, 0, 0);
        panel.add(colR, c);

        c.gridx = 0; c.gridy = 1;
        c.gridwidth = 2;
        c.weighty = 1.0;
        c.fill = GridBagConstraints.BOTH;
        panel.add(Box.createVerticalGlue(), c);

        return panel;
    }

    // =========================
    // Preferencias (CAMBIOS CÓNYUGE)
    // =========================
    private JComponent buildPreferenciasPanel() {
        JPanel panel = new JPanel(new GridBagLayout());
        panel.setBackground(Color.WHITE);
        panel.setBorder(new EmptyBorder(8, 8, 8, 8));

        JPanel colL = new JPanel(new GridBagLayout());
        colL.setBackground(Color.WHITE);
        JPanel colR = new JPanel(new GridBagLayout());
        colR.setBackground(Color.WHITE);

        cmbEstadoCivil = newCombo();
        cmbConyugeGrado = newCombo(); // queda, pero ya no lo mostramos si no querés (lo dejamos por compat.)
        cmbDestino1 = newCombo();
        cmbDestino2 = newCombo();

        txtObservEstadoCivil = new JTextField(18);

        // Cónyuge base
        txtConyugeApellido = new JTextField(18);
        txtConyugeNombre = new JTextField(18);
        spConyugeFechaNacimiento = newDatePicker();
        txtConyugeDNI = new JTextField(12);

        // Cónyuge militar + condicionados
        cmbConyugeEsMilitar = newCombo();
        txtConyugeNroId = new JTextField(12);
        cmbConyugeDestino = newCombo();

        cmbConyugeEscalafon = newCombo();
        cmbConyugeEspBasica = newCombo();
        cmbConyugeEspAvanzada = newCombo();

        // Hijos
        cmbHijos = newCombo();
        txtCantidadHijos = new JTextField(6);
        txtCantidadHijos.setEnabled(false);

        taImpedimentoTraslado = new JTextArea(6, 24);
        taImpedimentoTraslado.setLineWrap(true);
        taImpedimentoTraslado.setWrapStyleWord(true);
        JScrollPane spImpedimento = new JScrollPane(taImpedimentoTraslado);
        spImpedimento.setVerticalScrollBarPolicy(ScrollPaneConstants.VERTICAL_SCROLLBAR_AS_NEEDED);

        // Lado izquierdo: reordenado + condicional
        int rL = 0;
        addField(colL, rL++, "Estado civil", cmbEstadoCivil);
        addField(colL, rL++, "Observaciones estado civil", txtObservEstadoCivil);

        addSectionTitle(colL, rL++, "Datos del cónyuge");
        addField(colL, rL++, "Cónyuge - Apellido", txtConyugeApellido);
        addField(colL, rL++, "Cónyuge - Nombre", txtConyugeNombre);
        addField(colL, rL++, "Cónyuge - DNI", txtConyugeDNI);
        addField(colL, rL++, "Cónyuge - Fecha nacimiento", spConyugeFechaNacimiento);

        addField(colL, rL++, "¿Cónyuge es militar?", cmbConyugeEsMilitar);

        addField(colL, rL++, "Cónyuge - Grado", cmbConyugeGrado);
        addField(colL, rL++, "Cónyuge - Nro identificación", txtConyugeNroId);
        addField(colL, rL++, "Cónyuge - Destino", cmbConyugeDestino);

        addField(colL, rL++, "Cónyuge - Escalafón", cmbConyugeEscalafon);
        addField(colL, rL++, "Cónyuge - Especialidad básica / primaria", cmbConyugeEspBasica);
        addField(colL, rL++, "Cónyuge - Especialidad avanzada", cmbConyugeEspAvanzada);

        addField(colL, rL++, "Hijos (S/N)", cmbHijos);
        addField(colL, rL++, "Cantidad hijos", txtCantidadHijos);

        // Lado derecho (igual)
        int rR = 0;
        addField(colR, rR++, "Destino preferencia 1", cmbDestino1);
        addField(colR, rR++, "Destino preferencia 2", cmbDestino2);
        addField(colR, rR++, "Impedimento traslado (justificar)", spImpedimento);

        // Layout columnas
        GridBagConstraints c = new GridBagConstraints();
        c.gridx = 0; c.gridy = 0;
        c.weightx = 0.5; c.fill = GridBagConstraints.HORIZONTAL;
        c.anchor = GridBagConstraints.NORTHWEST;
        c.insets = new Insets(0, 0, 0, 12);
        panel.add(colL, c);

        c.gridx = 1; c.gridy = 0;
        c.insets = new Insets(0, 12, 0, 0);
        panel.add(colR, c);

        c.gridx = 0; c.gridy = 1;
        c.gridwidth = 2;
        c.weighty = 1.0;
        c.fill = GridBagConstraints.BOTH;
        panel.add(Box.createVerticalGlue(), c);

        // --- listeners cónyuge (evita la excepción setSelectedIndex out of bounds) ---
        setupConyugeUIStateAndListeners();
cmbHijos.addActionListener(e -> {
    boolean tieneHijos = "Sí".equalsIgnoreCase(valueOf(cmbHijos));
    txtCantidadHijos.setEnabled(tieneHijos);

    if (!tieneHijos) {
        txtCantidadHijos.setText("");
    }
});
        return panel;
    }

    private void setupConyugeUIStateAndListeners() {
        // Estado inicial: deshabilitados hasta que elija "Sí"
        applyConyugeMilitarUI(false);

        cmbConyugeEsMilitar.addActionListener(e -> {
            boolean esMil = "Sí".equalsIgnoreCase(valueOf(cmbConyugeEsMilitar));
            applyConyugeMilitarUI(esMil);
        });

        // cascada cónyuge (usa los mismos mapas)
        cmbConyugeEscalafon.addActionListener(e -> {
            if (!"Sí".equalsIgnoreCase(valueOf(cmbConyugeEsMilitar))) return;

            String escalafon = valueOf(cmbConyugeEscalafon);

            if (escalafon_especBasicaMap.isEmpty()) {
                cmbConyugeEspBasica.setEnabled(!escalafon.isEmpty());
                cmbConyugeEspAvanzada.setEnabled(!escalafon.isEmpty());
                return;
            }

            if (escalafon.isEmpty()) {
                cmbConyugeEspBasica.setEnabled(false);
                cmbConyugeEspAvanzada.setEnabled(false);
                fillComboSafe(cmbConyugeEspBasica, Collections.emptyList());
                fillComboSafe(cmbConyugeEspAvanzada, Collections.emptyList());
            } else {
                cmbConyugeEspBasica.setEnabled(true);
                Set<String> basicas = escalafon_especBasicaMap.getOrDefault(escalafon, new TreeSet<>(String.CASE_INSENSITIVE_ORDER));
                fillComboSafe(cmbConyugeEspBasica, new ArrayList<>(basicas));
                cmbConyugeEspAvanzada.setEnabled(false);
                fillComboSafe(cmbConyugeEspAvanzada, Collections.emptyList());
            }
        });

        cmbConyugeEspBasica.addActionListener(e -> {
            if (!"Sí".equalsIgnoreCase(valueOf(cmbConyugeEsMilitar))) return;

            String basica = valueOf(cmbConyugeEspBasica);

            if (especBasica_especAvanzadaMap.isEmpty()) {
                cmbConyugeEspAvanzada.setEnabled(!basica.isEmpty());
                return;
            }

            if (basica.isEmpty()) {
                cmbConyugeEspAvanzada.setEnabled(false);
                fillComboSafe(cmbConyugeEspAvanzada, Collections.emptyList());
            } else {
                cmbConyugeEspAvanzada.setEnabled(true);
                Set<String> avanzadas = especBasica_especAvanzadaMap.getOrDefault(basica, new TreeSet<>(String.CASE_INSENSITIVE_ORDER));
                fillComboSafe(cmbConyugeEspAvanzada, new ArrayList<>(avanzadas));
            }
        });
    }

    private void applyConyugeMilitarUI(boolean esMilitar) {
   txtConyugeNroId.setEnabled(esMilitar);
cmbConyugeDestino.setEnabled(esMilitar);

cmbConyugeGrado.setEnabled(esMilitar);
cmbConyugeEscalafon.setEnabled(esMilitar);

if (!esMilitar) {
    txtConyugeNroId.setText("");

    safeSelectIndex(cmbConyugeGrado, 0);
    safeSelectIndex(cmbConyugeDestino, 0);

    fillComboSafe(cmbConyugeEscalafon, getComboItemsOrEmpty(cmbEscalafon));
    safeSelectIndex(cmbConyugeEscalafon, 0);

    fillComboSafe(cmbConyugeEspBasica, Collections.emptyList());
    fillComboSafe(cmbConyugeEspAvanzada, Collections.emptyList());
    cmbConyugeEspBasica.setEnabled(false);
    cmbConyugeEspAvanzada.setEnabled(false);
} else {
    cmbConyugeEspBasica.setEnabled(false);
    cmbConyugeEspAvanzada.setEnabled(false);

    safeSelectIndex(cmbConyugeGrado, 0);
    safeSelectIndex(cmbConyugeDestino, 0);
    safeSelectIndex(cmbConyugeEscalafon, 0);

    fillComboSafe(cmbConyugeEspBasica, Collections.emptyList());
    fillComboSafe(cmbConyugeEspAvanzada, Collections.emptyList());
}
    }

    private List<String> getComboItemsOrEmpty(JComboBox<String> cb) {
        if (cb == null) return Collections.emptyList();
        List<String> out = new ArrayList<>();
        for (int i = 0; i < cb.getItemCount(); i++) {
            String s = cb.getItemAt(i);
            if (s != null && !s.isBlank()) out.add(s);
        }
        return out;
    }

    // =========================
    // Footer
    // =========================
    private JComponent buildFooterButtons() {
        JPanel footer = new JPanel(new BorderLayout());
        footer.setBackground(new Color(0xEAF3FF));

        JPanel right = new JPanel(new FlowLayout(FlowLayout.RIGHT, 10, 0));
        right.setBackground(new Color(0xEAF3FF));

        JButton btnLimpiar = new JButton("Limpiar");
        JButton btnGuardar = new JButton("Guardar");
        JButton btnCerrar = new JButton("Cerrar");

        styleButtonOutline(btnLimpiar, FORM_BUTTON_PRIMARY);
        styleButtonPrimary(btnGuardar, FORM_BUTTON_PRIMARY);
        styleButtonSecondary(btnCerrar, FORM_BUTTON_SECONDARY);

        btnLimpiar.addActionListener(this::onLimpiar);
        btnGuardar.addActionListener(this::onGuardar);
        btnCerrar.addActionListener(e -> dispose());

        right.add(btnLimpiar);
        right.add(btnGuardar);
        right.add(btnCerrar);

        footer.add(right, BorderLayout.EAST);
        return footer;
    }

    // =========================
    // Carga combos (agrega cónyuge militar + cascada)
    // =========================
    private void loadComboData() {
        fillComboSafe(cmbGrado, listProvider.grados());
        fillComboSafe(cmbProvincia, listProvider.provincias());
        fillComboSafe(cmbEstadoCivil, listProvider.estadosCiviles());
        fillComboSafe(cmbConyugeGrado, listProvider.grados());

        // Sí/No
       List<String> siNo = Arrays.asList("", "Sí", "No");
fillComboSafe(cmbHijos, siNo);

List<String> gruposSanguineos = Arrays.asList(
        "A+", "A-",
        "B+", "B-",
        "AB+", "AB-",
        "O+", "O-"
);
fillComboSafe(cmbFactorSanguineo, gruposSanguineos);
        if (cmbConyugeEsMilitar != null) fillComboSafe(cmbConyugeEsMilitar, siNo);

        // ====== Escalafón/Especialidades DESDE EXCEL (resources) ======
        EspecialidadesData espData = readEspecialidadesDesdeExcelResource(
                "/ESPECIALIDADES DEL PERSONAL MILITAR SUBALTERNO (EN PROCESO).xlsx",
                0,
                1,   // B (0-based)
                2,   // C
                4,   // E
                7
        );

        if (espData != null && !espData.escalafones.isEmpty()) {
            fillComboSafe(cmbEscalafon, espData.escalafones);
            fillComboSafe(cmbEspBasica, espData.basicas);
            fillComboSafe(cmbEspAvanzada, espData.avanzadas);

            escalafon_especBasicaMap.clear();
            escalafon_especBasicaMap.putAll(espData.mapEscalafonToBasicas);

            especBasica_especAvanzadaMap.clear();
            especBasica_especAvanzadaMap.putAll(espData.mapBasicaToAvanzadas);
        } else {
            fillComboSafe(cmbEspBasica, listProvider.espBasica());
            fillComboSafe(cmbEspAvanzada, listProvider.espAvanzada());
            fillComboSafe(cmbEscalafon, listProvider.escalafon());

            escalafon_especBasicaMap.clear();
            especBasica_especAvanzadaMap.clear();
        }

        // combos cónyuge especialidad (mismas listas)
        if (cmbConyugeEscalafon != null) fillComboSafe(cmbConyugeEscalafon, getComboItemsOrEmpty(cmbEscalafon));
        if (cmbConyugeEspBasica != null) fillComboSafe(cmbConyugeEspBasica, Collections.emptyList());
        if (cmbConyugeEspAvanzada != null) fillComboSafe(cmbConyugeEspAvanzada, Collections.emptyList());

        // Destinos desde CSV resources
        List<String> destinos = readDestinos();
fillComboSafe(cmbDestino1, destinos);
fillComboSafe(cmbDestino2, destinos);
fillComboSafe(cmbUnidadRevista, destinos);
fillComboSafe(cmbDestinoAnterior, destinos);
fillComboSafe(cmbConyugeDestino, destinos);

        fillComboSafe(cmbSeDesempena, siNo);
        fillComboSafe(cmbCumpleTurno, siNo);
        fillComboSafe(cmbPoseeAptoFisico, siNo);
        fillComboSafe(cmbDeseaPermanecer, siNo);
        fillComboSafe(cmbRindioSidiel, siNo);
        fillComboSafe(cmbActividadProfesor, siNo);
        fillComboSafe(cmbDictaActualmente, siNo);
        fillComboSafe(cmbCumplioCampanasAntarticas, siNo);
    }

    private static class EspecialidadesData {
        final List<String> escalafones = new ArrayList<>();
        final List<String> basicas = new ArrayList<>();
        final List<String> avanzadas = new ArrayList<>();
        final Map<String, Set<String>> mapEscalafonToBasicas = new HashMap<>();
        final Map<String, Set<String>> mapBasicaToAvanzadas = new HashMap<>();
    }

    private EspecialidadesData readEspecialidadesDesdeExcelResource(
            String resourcePath,
            int sheetIndex,
            int colEscalafon,
            int colBasica,
            int colAvanzada,
            int startRow
    ) {
        try (InputStream is = FichaDatosPersonalesFrame.class.getResourceAsStream(resourcePath)) {
            if (is == null) return null;

            EspecialidadesData data = new EspecialidadesData();

            try (Workbook wb = new XSSFWorkbook(is)) {
                Sheet sh = wb.getNumberOfSheets() > sheetIndex ? wb.getSheetAt(sheetIndex) : wb.getSheetAt(0);
                if (sh == null) return null;

                Set<String> escalafonesSet = new TreeSet<>(String.CASE_INSENSITIVE_ORDER);
                Set<String> basicasSet = new TreeSet<>(String.CASE_INSENSITIVE_ORDER);
                Set<String> avanzadasSet = new TreeSet<>(String.CASE_INSENSITIVE_ORDER);

                int lastRow = sh.getLastRowNum();
                String lastEsc = "";
                String lastBas = "";

                for (int r = Math.max(startRow, sh.getFirstRowNum()); r <= lastRow; r++) {
                    Row row = sh.getRow(r);
                    if (row == null) continue;

                    String escRaw = getCellString(row.getCell(colEscalafon));
                    String basRaw = getCellString(row.getCell(colBasica));
                    String avzRaw = getCellString(row.getCell(colAvanzada));

                    String esc = !escRaw.isEmpty() ? escRaw : lastEsc;
                    String bas = !basRaw.isEmpty() ? basRaw : lastBas;
                    String avz = avzRaw;

                    if (!escRaw.isEmpty()) {
                        lastEsc = escRaw;
                        lastBas = "";
                    }
                    if (!basRaw.isEmpty()) lastBas = basRaw;

                    if (esc.isEmpty() && bas.isEmpty() && avz.isEmpty()) continue;

                    String escLow = esc.toLowerCase(Locale.ROOT);
                    String basLow = bas.toLowerCase(Locale.ROOT);
                    if (escLow.contains("escalaf") || basLow.contains("especialidad")) continue;

                    if (!esc.isEmpty()) escalafonesSet.add(esc);
                    if (!bas.isEmpty()) basicasSet.add(bas);
                    if (!avz.isEmpty()) avanzadasSet.add(avz);

                    if (!esc.isEmpty() && !bas.isEmpty()) {
                        data.mapEscalafonToBasicas
                                .computeIfAbsent(esc, k -> new TreeSet<>(String.CASE_INSENSITIVE_ORDER))
                                .add(bas);
                    }
                    if (!bas.isEmpty() && !avz.isEmpty()) {
                        data.mapBasicaToAvanzadas
                                .computeIfAbsent(bas, k -> new TreeSet<>(String.CASE_INSENSITIVE_ORDER))
                                .add(avz);
                    }
                }

                data.escalafones.addAll(escalafonesSet);
                data.basicas.addAll(basicasSet);
                data.avanzadas.addAll(avanzadasSet);

                return data;
            }
        } catch (Exception ex) {
            return null;
        }
    }

    private static String getCellString(Cell c) {
        if (c == null) return "";
        try {
            CellType t = c.getCellType();
            if (t == CellType.FORMULA) {
                try {
                    return c.getStringCellValue().trim();
                } catch (Exception ignored) {
                    try {
                        double d = c.getNumericCellValue();
                        if (Math.floor(d) == d) return String.valueOf((long) d);
                        return String.valueOf(d);
                    } catch (Exception ignored2) {
                        return "";
                    }
                }
            }

            switch (t) {
                case STRING:
                    return c.getStringCellValue().trim();
                case NUMERIC:
                    double d = c.getNumericCellValue();
                    if (Math.floor(d) == d) return String.valueOf((long) d);
                    return String.valueOf(d);
                case BOOLEAN:
                    return String.valueOf(c.getBooleanCellValue());
                default:
                    return "";
            }
        } catch (Exception e) {
            return "";
        }
    }

    private static List<String> readDestinos() {
        List<String> out = new ArrayList<>();
        try (InputStream is = FichaDatosPersonalesFrame.class.getResourceAsStream("/destinos.csv")) {
            if (is != null) {
                try (BufferedReader br = new BufferedReader(new InputStreamReader(is, StandardCharsets.UTF_8))) {
                    String line;
                    while ((line = br.readLine()) != null) {
                        line = line.trim();
                        if (!line.isEmpty()) out.add(line);
                    }
                }
            }
        } catch (Exception ignored) {}
        if (out.isEmpty()) out.addAll(Arrays.asList("Base A", "Base B", "Base C"));
        return out;
    }

    // =========================
    // Combo helpers (anti crash)
    // =========================
    private void fillComboSafe(JComboBox<String> combo, List<String> items) {
        if (combo == null) return;
        combo.removeAllItems();
        combo.addItem("");
        if (items != null) {
            for (String s : items) combo.addItem(s);
        }
        safeSelectIndex(combo, 0);
    }

    private void safeSelectIndex(JComboBox<?> combo, int idx) {
        if (combo == null) return;
        int n = combo.getItemCount();
        if (n <= 0) return;
        int safe = Math.max(0, Math.min(idx, n - 1));
        combo.setSelectedIndex(safe);
    }

    private void setupCascadeListeners() {
        cmbEscalafon.addActionListener(e -> {
            String escalafon = valueOf(cmbEscalafon);

            if (escalafon_especBasicaMap.isEmpty()) {
                cmbEspBasica.setEnabled(!escalafon.isEmpty());
                cmbEspAvanzada.setEnabled(!escalafon.isEmpty());
                return;
            }

            if (escalafon.isEmpty()) {
                cmbEspBasica.setEnabled(false);
                cmbEspAvanzada.setEnabled(false);
                fillComboSafe(cmbEspBasica, Collections.emptyList());
                fillComboSafe(cmbEspAvanzada, Collections.emptyList());
            } else {
                cmbEspBasica.setEnabled(true);
                Set<String> basicas = escalafon_especBasicaMap.getOrDefault(escalafon, new TreeSet<>(String.CASE_INSENSITIVE_ORDER));
                fillComboSafe(cmbEspBasica, new ArrayList<>(basicas));
                cmbEspAvanzada.setEnabled(false);
                fillComboSafe(cmbEspAvanzada, Collections.emptyList());
            }
        });

        cmbEspBasica.addActionListener(e -> {
            String basica = valueOf(cmbEspBasica);

            if (especBasica_especAvanzadaMap.isEmpty()) {
                cmbEspAvanzada.setEnabled(!basica.isEmpty());
                return;
            }

            if (basica.isEmpty()) {
                cmbEspAvanzada.setEnabled(false);
                fillComboSafe(cmbEspAvanzada, Collections.emptyList());
            } else {
                cmbEspAvanzada.setEnabled(true);
                Set<String> avanzadas = especBasica_especAvanzadaMap.getOrDefault(basica, new TreeSet<>(String.CASE_INSENSITIVE_ORDER));
                fillComboSafe(cmbEspAvanzada, new ArrayList<>(avanzadas));
            }
        });
    }

    // =========================
    // Guardar / Limpiar
    // =========================
    private void onLimpiar(ActionEvent e) {
        forEachTextField(
                () -> txtIOSFA, () -> txtDNI, () -> txtApellido, () -> txtNombres,
                () -> txtLugarNac, () -> txtCUIL, () -> txtCBU, () -> txtEmailInst, () -> txtCelular,
                () -> txtUsuarioGDE, () -> txtRTI,
                () -> txtDomicilioCalle, () -> txtNumeroCalle, () -> txtLocalidad, () -> txtCP,
                () -> txtObservEstadoCivil,
                () -> txtConyugeNroId, () -> txtConyugeApellido, () -> txtConyugeNombre,
                () -> txtConyugeDNI,
                () -> txtCantidadHijos,
                () -> txtFuncion, () -> txtPromedioTurnos,
                () -> txtIdioma1, () -> txtNivelIdioma1, () -> txtIdioma2, () -> txtNivelIdioma2, () -> txtIdioma3, () -> txtNivelIdioma3,
                () -> txtCapInstTitulo, () -> txtCapInstExpedidoPor, () -> txtMaxCapExtraTitulo, () -> txtMaxCapExtraExpedidoPor,
                () -> txtAsignaturaTemas, () -> txtModalidad, () -> txtTituloHabilitante,
                () -> txtComisionExteriorMotivo, () -> txtComisionExteriorPaisCiudad, () -> txtCantidadCampanas, () -> txtDotacionGpoTareas, () -> txtCargoDesempenado,
                () -> txtAniosEnGrado, () -> txtAniosEnEspecialidad,
                () -> txtDestinoInterno, () -> txtCargo
        ).forEach(tf -> { if (tf != null) tf.setText(""); });

        if (taImpedimentoTraslado != null) taImpedimentoTraslado.setText("");

        forEachCombo(
                () -> cmbGrado, () -> cmbEspBasica, () -> cmbEspAvanzada, () -> cmbEscalafon,() -> cmbFactorSanguineo,
                () -> cmbProvincia, () -> cmbEstadoCivil, () -> cmbDestino1, () -> cmbDestino2,
                ()-> cmbHijos,
                () -> cmbSeDesempena, () -> cmbCumpleTurno, () -> cmbPoseeAptoFisico, () -> cmbDeseaPermanecer,
                () -> cmbRindioSidiel, () -> cmbActividadProfesor, () -> cmbDictaActualmente, () -> cmbCumplioCampanasAntarticas,
                () -> cmbUnidadRevista, () -> cmbDestinoAnterior,
               () -> cmbConyugeEsMilitar, () -> cmbConyugeGrado, () -> cmbConyugeDestino,
() -> cmbConyugeEscalafon, () -> cmbConyugeEspBasica, () -> cmbConyugeEspAvanzada
        ).forEach(cb -> safeSelectIndex(cb, 0));

        forEachDatePicker(
                () -> spFechaNacimiento, () -> spFechaIngreso, () -> spAptitudPsicofisicaFecha, () -> spFechaCondicionTiro,
                () -> spFechaNivel1, () -> spFechaNivel2, () -> spFechaNivel3, () -> spCapInstFecha, () -> spMaxCapExtraFecha,
                () -> spFechaInicioComision, () -> spFechaFinComision, () -> spPeriodoDesde, () -> spPeriodoHasta,
                () -> spConyugeFechaNacimiento
        ).forEach(this::clearDatePicker);

        applyConyugeMilitarUI(false);
    }

    private void onGuardar(ActionEvent e) {
        if (!validateForm()) {
            return;
        }

        model.iosfa = txtIOSFA.getText().trim();
        model.dni = txtDNI.getText().trim();
        model.grado = valueOf(cmbGrado);
        model.apellido = txtApellido.getText().trim();
        model.nombres = txtNombres.getText().trim();
        model.lugarNac = txtLugarNac.getText().trim();

        model.fechaNacimiento = dateTextOf(spFechaNacimiento);
        model.fechaIngreso = dateTextOf(spFechaIngreso);
        model.aniosEnGrado = txtAniosEnGrado.getText().trim();
        model.aniosEnEspecialidad = txtAniosEnEspecialidad.getText().trim();
        model.seDesempena = valueOf(cmbSeDesempena);
        model.cumpleTurno = valueOf(cmbCumpleTurno);
        model.funcion = txtFuncion.getText().trim();
        model.promedioTurnos = txtPromedioTurnos.getText().trim();
        model.aptitudPsicofisicaFecha = dateTextOf(spAptitudPsicofisicaFecha);
        model.fechaCondicionTiro = dateTextOf(spFechaCondicionTiro);
        model.poseeAptoFisico = valueOf(cmbPoseeAptoFisico);

        model.escalafon = valueOf(cmbEscalafon);
        model.espBasica = valueOf(cmbEspBasica);
        model.espAvanzada = valueOf(cmbEspAvanzada);

        model.cuil = txtCUIL.getText().trim();
        model.cbu = txtCBU.getText().trim();
        model.emailInst = txtEmailInst.getText().trim();
        model.celular = txtCelular.getText().trim();
        model.usuarioGDE = txtUsuarioGDE.getText().trim();
        model.rti = txtRTI.getText().trim();
        model.factorSanguineo = valueOf(cmbFactorSanguineo);
        model.unidadRevista = valueOf(cmbUnidadRevista);
        model.destinoInterno = txtDestinoInterno.getText().trim();
        model.cargo = txtCargo.getText().trim();
        model.destinoAnterior = valueOf(cmbDestinoAnterior);
        model.deseaPermanecer = valueOf(cmbDeseaPermanecer);

        model.domicilioCalle = txtDomicilioCalle.getText().trim();
        model.domicilioNumero = txtNumeroCalle.getText().trim();
        model.localidad = txtLocalidad.getText().trim();
        model.cp = txtCP.getText().trim();
        model.provincia = valueOf(cmbProvincia);

        model.estadoCivil = valueOf(cmbEstadoCivil);
        model.destino1 = valueOf(cmbDestino1);
        model.destino2 = valueOf(cmbDestino2);
        model.observacionesEstadoCivil = txtObservEstadoCivil.getText().trim();

        // Cónyuge base
        model.conyugeApellido = txtConyugeApellido.getText().trim();
        model.conyugeNombre = txtConyugeNombre.getText().trim();
        model.conyugeFechaNacimiento = dateTextOf(spConyugeFechaNacimiento);
        model.conyugeDNI = txtConyugeDNI.getText().trim();

        // Cónyuge militar + campos condicionados
        model.conyugeEsMilitar = valueOf(cmbConyugeEsMilitar);
        boolean esMil = "Sí".equalsIgnoreCase(model.conyugeEsMilitar);

        model.conyugeGrado = valueOf(cmbConyugeGrado);
        model.conyugeNroIdentificacion = esMil ? txtConyugeNroId.getText().trim() : "";
        model.conyugeDestino = esMil ? valueOf(cmbConyugeDestino) : "";

        String cEsc = esMil ? valueOf(cmbConyugeEscalafon) : "";
        String cBas = esMil ? valueOf(cmbConyugeEspBasica) : "";
        String cAvz = esMil ? valueOf(cmbConyugeEspAvanzada) : "";
        model.conyugeEspecialidad = esMil ? buildConyugeEspecialidadString(cEsc, cBas, cAvz) : "";

        model.hijos = valueOf(cmbHijos);
        model.cantidadHijos = txtCantidadHijos.getText().trim();
        model.impedimentoTraslado = (taImpedimentoTraslado == null) ? "" : taImpedimentoTraslado.getText().trim();

        model.idioma1 = txtIdioma1.getText().trim();
        model.nivelIdioma1 = txtNivelIdioma1.getText().trim();
        model.fechaNivel1 = dateTextOf(spFechaNivel1);
        model.idioma2 = txtIdioma2.getText().trim();
        model.nivelIdioma2 = txtNivelIdioma2.getText().trim();
        model.fechaNivel2 = dateTextOf(spFechaNivel2);
        model.idioma3 = txtIdioma3.getText().trim();
        model.nivelIdioma3 = txtNivelIdioma3.getText().trim();
        model.fechaNivel3 = dateTextOf(spFechaNivel3);
        model.rindioSidiel = valueOf(cmbRindioSidiel);

        model.capInstTitulo = txtCapInstTitulo.getText().trim();
        model.capInstExpedidoPor = txtCapInstExpedidoPor.getText().trim();
        model.capInstFecha = dateTextOf(spCapInstFecha);
        model.maxCapExtraTitulo = txtMaxCapExtraTitulo.getText().trim();
        model.maxCapExtraExpedidoPor = txtMaxCapExtraExpedidoPor.getText().trim();
        model.maxCapExtraFecha = dateTextOf(spMaxCapExtraFecha);
        model.actividadProfesor = valueOf(cmbActividadProfesor);
        model.asignaturaTemas = txtAsignaturaTemas.getText().trim();
        model.dictaActualmente = valueOf(cmbDictaActualmente);
        model.modalidad = txtModalidad.getText().trim();
        model.tituloHabilitante = txtTituloHabilitante.getText().trim();

        model.comisionExteriorMotivo = txtComisionExteriorMotivo.getText().trim();
        model.comisionExteriorPaisCiudad = txtComisionExteriorPaisCiudad.getText().trim();
        model.fechaInicioComision = dateTextOf(spFechaInicioComision);
        model.fechaFinComision = dateTextOf(spFechaFinComision);
        model.cumplioCampanasAntarticas = valueOf(cmbCumplioCampanasAntarticas);
        model.cantidadCampanas = txtCantidadCampanas.getText().trim();
        model.dotacionGpoTareas = txtDotacionGpoTareas.getText().trim();
        model.cargoDesempenado = txtCargoDesempenado.getText().trim();
        model.periodoDesde = dateTextOf(spPeriodoDesde);
        model.periodoHasta = dateTextOf(spPeriodoHasta);

        try {
            String fileName = saveToExcel();
            JOptionPane.showMessageDialog(this, "Datos guardados en: " + fileName, "Guardar", JOptionPane.INFORMATION_MESSAGE);
        } catch (IOException ex) {
            JOptionPane.showMessageDialog(this, "Error al guardar: " + ex.getMessage(), "Error", JOptionPane.ERROR_MESSAGE);
        }
    }

    private boolean validateForm() {
        if (!validateRequiredText(txtIOSFA, "IOSFA")) return false;
        if (!validateNoDots(txtIOSFA, "IOSFA")) return false;

        if (!validateRequiredText(txtDNI, "DNI")) return false;
        if (!validateNoDots(txtDNI, "DNI")) return false;

        if (!validateRequiredText(txtApellido, "Apellido")) return false;
        if (!validateRequiredText(txtNombres, "Nombres")) return false;
        if (!validateRequiredText(txtLugarNac, "Lugar de nacimiento")) return false;

        if (!validateDatePicker(spFechaNacimiento, "Fecha de nacimiento")) return false;
        if (!validateDatePicker(spFechaIngreso, "Fecha de ingreso")) return false;
        if (!validateDatePicker(spAptitudPsicofisicaFecha, "Fecha aptitud psicofísica")) return false;
        if (!validateDatePicker(spFechaCondicionTiro, "Fecha condición tiro")) return false;

        if (!validateDigitsOnly(txtAniosEnGrado, "Años en el grado")) return false;
        if (!validateDigitsOnly(txtAniosEnEspecialidad, "Años en la especialidad")) return false;
        if (!validateDigitsOnly(txtPromedioTurnos, "Promedio de turnos")) return false;

        if (!validateRequiredText(txtCUIL, "CUIL")) return false;
        if (!validateRequiredText(txtCBU, "CBU")) return false;
        if (!validateDigitsOnly(txtCBU, "CBU")) return false;
        if (!validateRequiredText(txtCelular, "Celular")) return false;
        if (!validateDigitsOnly(txtCelular, "Celular")) return false;

        if (!validateRequiredText(txtEmailInst, "Email institucional")) return false;
        if (!validateRequiredText(txtUsuarioGDE, "Usuario GDE")) return false;
        if (!validateRequiredText(txtRTI, "RTI")) return false;

        if (!validateRequiredText(txtDestinoInterno, "Destino interno")) return false;
        if (!validateRequiredText(txtCargo, "Cargo")) return false;
        if (!validateRequiredText(txtDomicilioCalle, "Domicilio calle")) return false;
        if (!validateRequiredText(txtNumeroCalle, "Número de domicilio")) return false;
        if (!validateRequiredText(txtLocalidad, "Localidad")) return false;
        if (!validateRequiredText(txtCP, "Código postal")) return false;

        if (!validateCombo(cmbGrado, "Grado")) return false;
        if (!validateCombo(cmbSeDesempena, "¿Se desempeña en la especialidad actualmente?") ) return false;
        if (!validateCombo(cmbCumpleTurno, "¿Cumple turno o servicio?") ) return false;
        if (!validateCombo(cmbPoseeAptoFisico, "¿Posee apto físico?") ) return false;
        if (!validateCombo(cmbDeseaPermanecer, "Desea permanecer en el destino actual?") ) return false;
        if (!validateCombo(cmbFactorSanguineo, "Factor sanguíneo")) return false;
        if (!validateCombo(cmbUnidadRevista, "Unidad revista")) return false;
        if (!validateCombo(cmbDestinoAnterior, "Destino anterior")) return false;
        if (!validateCombo(cmbProvincia, "Provincia")) return false;
        if (!validateCombo(cmbEstadoCivil, "Estado civil")) return false;

        if (!validateCombo(cmbHijos, "¿Hijos?")) return false;
        if ("Sí".equalsIgnoreCase(valueOf(cmbHijos)) && !validateDigitsOnly(txtCantidadHijos, "Cantidad de hijos")) return false;

        if (!validateCombo(cmbConyugeEsMilitar, "¿Cónyuge es militar?")) return false;
        if ("Sí".equalsIgnoreCase(valueOf(cmbConyugeEsMilitar))) {
            if (!validateRequiredText(txtConyugeNroId, "Número de identificación del cónyuge")) return false;
            if (!validateCombo(cmbConyugeDestino, "Destino del cónyuge")) return false;
            if (!validateCombo(cmbConyugeEscalafon, "Escalafón del cónyuge")) return false;
            if (!validateCombo(cmbConyugeEspBasica, "Especialidad básica del cónyuge")) return false;
            if (!validateCombo(cmbConyugeEspAvanzada, "Especialidad avanzada del cónyuge")) return false;
        }

        return true;
    }

    private boolean validateRequiredText(JTextField field, String fieldName) {
        if (field == null || field.getText().trim().isEmpty()) {
            showValidationMessage(fieldName + " es obligatorio.");
            if (field != null) field.requestFocusInWindow();
            return false;
        }
        return true;
    }

    private boolean validateNoDots(JTextField field, String fieldName) {
        if (field == null) return true;
        if (field.getText().trim().contains(".")) {
            showValidationMessage(fieldName + " no puede contener puntos.");
            field.requestFocusInWindow();
            return false;
        }
        return true;
    }

    private boolean validateDatePicker(JDateChooser datePicker, String fieldName) {
        if (datePicker == null || datePicker.getDate() == null) {
            showValidationMessage(fieldName + " es obligatorio.");
            if (datePicker != null) datePicker.requestFocusInWindow();
            return false;
        }
        return true;
    }

    private boolean validateDigitsOnly(JTextField field, String fieldName) {
        if (field == null || field.getText().trim().isEmpty()) {
            showValidationMessage(fieldName + " es obligatorio.");
            if (field != null) field.requestFocusInWindow();
            return false;
        }
        if (!field.getText().trim().matches("\\d+")) {
            showValidationMessage(fieldName + " solo debe contener números.");
            field.requestFocusInWindow();
            return false;
        }
        return true;
    }

    private boolean validateCombo(JComboBox<String> combo, String fieldName) {
        if (combo == null || valueOf(combo).isEmpty()) {
            showValidationMessage(fieldName + " debe seleccionarse.");
            if (combo != null) combo.requestFocusInWindow();
            return false;
        }
        return true;
    }

    private void showValidationMessage(String message) {
        JOptionPane.showMessageDialog(this, message, "Validación", JOptionPane.WARNING_MESSAGE);
    }

    private String buildConyugeEspecialidadString(String esc, String bas, String avz) {
        // Formato simple para guardar en una sola columna existente:
        // "Escalafón | Básica | Avanzada" (omitimos vacíos)
        List<String> parts = new ArrayList<>();
        if (esc != null && !esc.isBlank()) parts.add(esc.trim());
        if (bas != null && !bas.isBlank()) parts.add(bas.trim());
        if (avz != null && !avz.isBlank()) parts.add(avz.trim());
        return String.join(" | ", parts);
    }

    private String saveToExcel() throws IOException {
        XSSFWorkbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("Ficha Datos");
        CellStyle textStyle = workbook.createCellStyle();
        textStyle.setDataFormat(workbook.createDataFormat().getFormat("@"));

        Row headerRow = sheet.createRow(0);
        String[] headers = {
                "IOSFA", "DNI", "Grado", "Apellido", "Nombres", "Lugar Nacimiento",
                "Fecha de nacimiento", "Fecha de ingreso", "Años en el grado", "Años en la especialidad",
                "Se desempeña", "Cumple turno", "Escalafón", "Especialidad Básica", "Especialidad Avanzada",
                "Función", "Promedio turnos", "Aptitud psicofísica (fecha)", "Condición tiro (fecha)", "Apto físico",
                "CUIL", "CBU", "Email Inst", "Celular", "Usuario GDE", "RTI", "Factor sanguíneo",
                "Unidad revista", "Destino interno", "Cargo", "Destino anterior", "Desea permanecer",
                "Domicilio Calle", "Domicilio Número", "Localidad", "CP", "Provincia",
                "Estado Civil", "Conyuge Grado", "Destino 1", "Destino 2",
                "Obs Estado Civil", "Conyuge Nro ID", "Conyuge Apellido", "Conyuge Nombre",
                "Conyuge Fecha Nac", "Conyuge DNI", "Conyuge Especialidad", "Conyuge Destino",
                "Hijos", "Cantidad Hijos", "Impedimento Traslado",
                "Idioma 1", "Nivel 1", "Fecha nivel 1",
                "Idioma 2", "Nivel 2", "Fecha nivel 2",
                "Idioma 3", "Nivel 3", "Fecha nivel 3",
                "Rindió SIDIEL",
                "Capacitación Institucional Título", "Capacitación Institucional Expedido Por", "Capacitación Institucional Fecha",
                "Máxima Capacitación Extra Institucional Título", "Maxima Capacitación Extra Institucional Expedido Por", "Maxima Capacitación Extra Institucional Fecha",
                "Profesor", "Asignatura/Temas", "Dicta actualmente", "Modalidad", "Título habilitante",
                "Comisión Motivo", "Comisión País/Ciudad", "Inicio comisión", "Fin comisión",
                "Campañas antárticas", "Cantidad campañas", "Dotación/GPO", "Cargo desempeñado",
                "Periodo desde", "Periodo hasta"
        };

        for (int i = 0; i < headers.length; i++) {
            Cell cell = headerRow.createCell(i);
            cell.setCellValue(headers[i]);
        }

        Row row = sheet.createRow(1);
        int ci = 0;

        row.createCell(ci++).setCellValue(model.iosfa);
        row.createCell(ci++).setCellValue(model.dni);
        row.createCell(ci++).setCellValue(model.grado);
        row.createCell(ci++).setCellValue(model.apellido);
        row.createCell(ci++).setCellValue(model.nombres);
        row.createCell(ci++).setCellValue(model.lugarNac);

        row.createCell(ci++).setCellValue(model.fechaNacimiento);
        row.createCell(ci++).setCellValue(model.fechaIngreso);
        row.createCell(ci++).setCellValue(model.aniosEnGrado);
        row.createCell(ci++).setCellValue(model.aniosEnEspecialidad);

        row.createCell(ci++).setCellValue(model.seDesempena);
        row.createCell(ci++).setCellValue(model.cumpleTurno);
        row.createCell(ci++).setCellValue(model.escalafon);
        row.createCell(ci++).setCellValue(model.espBasica);
        row.createCell(ci++).setCellValue(model.espAvanzada);

        row.createCell(ci++).setCellValue(model.funcion);
        row.createCell(ci++).setCellValue(model.promedioTurnos);
        row.createCell(ci++).setCellValue(model.aptitudPsicofisicaFecha);
        row.createCell(ci++).setCellValue(model.fechaCondicionTiro);
        row.createCell(ci++).setCellValue(model.poseeAptoFisico);

        row.createCell(ci++).setCellValue(model.cuil);
        Cell cbuCell = row.createCell(ci++, CellType.STRING);
        cbuCell.setCellValue(model.cbu);
        cbuCell.setCellStyle(textStyle);
        row.createCell(ci++).setCellValue(model.emailInst);
        Cell celularCell = row.createCell(ci++, CellType.STRING);
        celularCell.setCellValue(model.celular);
        celularCell.setCellStyle(textStyle);
        row.createCell(ci++).setCellValue(model.usuarioGDE);
        row.createCell(ci++).setCellValue(model.rti);
        row.createCell(ci++).setCellValue(model.factorSanguineo);

        row.createCell(ci++).setCellValue(model.unidadRevista);
        row.createCell(ci++).setCellValue(model.destinoInterno);
        row.createCell(ci++).setCellValue(model.cargo);
        row.createCell(ci++).setCellValue(model.destinoAnterior);
        row.createCell(ci++).setCellValue(model.deseaPermanecer);

        Cell domicilioCalleCell = row.createCell(ci++, CellType.STRING);
        domicilioCalleCell.setCellValue(model.domicilioCalle);
        domicilioCalleCell.setCellStyle(textStyle);
        Cell domicilioNumeroCell = row.createCell(ci++, CellType.STRING);
        domicilioNumeroCell.setCellValue(model.domicilioNumero);
        domicilioNumeroCell.setCellStyle(textStyle);
        Cell localidadCell = row.createCell(ci++, CellType.STRING);
        localidadCell.setCellValue(model.localidad);
        localidadCell.setCellStyle(textStyle);
        Cell cpCell = row.createCell(ci++, CellType.STRING);
        cpCell.setCellValue(model.cp);
        cpCell.setCellStyle(textStyle);
        row.createCell(ci++).setCellValue(model.provincia);

        row.createCell(ci++).setCellValue(model.estadoCivil);
        row.createCell(ci++).setCellValue(model.conyugeGrado);
        row.createCell(ci++).setCellValue(model.destino1);
        row.createCell(ci++).setCellValue(model.destino2);

        row.createCell(ci++).setCellValue(model.observacionesEstadoCivil);
        row.createCell(ci++).setCellValue(model.conyugeNroIdentificacion);
        row.createCell(ci++).setCellValue(model.conyugeApellido);
        row.createCell(ci++).setCellValue(model.conyugeNombre);
        row.createCell(ci++).setCellValue(model.conyugeFechaNacimiento);
        row.createCell(ci++).setCellValue(model.conyugeDNI);
        row.createCell(ci++).setCellValue(model.conyugeEspecialidad);
        row.createCell(ci++).setCellValue(model.conyugeDestino);

        row.createCell(ci++).setCellValue(model.hijos);
        row.createCell(ci++).setCellValue(model.cantidadHijos);
        row.createCell(ci++).setCellValue(model.impedimentoTraslado);

        row.createCell(ci++).setCellValue(model.idioma1);
        row.createCell(ci++).setCellValue(model.nivelIdioma1);
        row.createCell(ci++).setCellValue(model.fechaNivel1);
        row.createCell(ci++).setCellValue(model.idioma2);
        row.createCell(ci++).setCellValue(model.nivelIdioma2);
        row.createCell(ci++).setCellValue(model.fechaNivel2);
        row.createCell(ci++).setCellValue(model.idioma3);
        row.createCell(ci++).setCellValue(model.nivelIdioma3);
        row.createCell(ci++).setCellValue(model.fechaNivel3);
        row.createCell(ci++).setCellValue(model.rindioSidiel);

        row.createCell(ci++).setCellValue(model.capInstTitulo);
        row.createCell(ci++).setCellValue(model.capInstExpedidoPor);
        row.createCell(ci++).setCellValue(model.capInstFecha);
        row.createCell(ci++).setCellValue(model.maxCapExtraTitulo);
        row.createCell(ci++).setCellValue(model.maxCapExtraExpedidoPor);
        row.createCell(ci++).setCellValue(model.maxCapExtraFecha);

        row.createCell(ci++).setCellValue(model.actividadProfesor);
        row.createCell(ci++).setCellValue(model.asignaturaTemas);
        row.createCell(ci++).setCellValue(model.dictaActualmente);
        row.createCell(ci++).setCellValue(model.modalidad);
        row.createCell(ci++).setCellValue(model.tituloHabilitante);

        row.createCell(ci++).setCellValue(model.comisionExteriorMotivo);
        row.createCell(ci++).setCellValue(model.comisionExteriorPaisCiudad);
        row.createCell(ci++).setCellValue(model.fechaInicioComision);
        row.createCell(ci++).setCellValue(model.fechaFinComision);

        row.createCell(ci++).setCellValue(model.cumplioCampanasAntarticas);
        row.createCell(ci++).setCellValue(model.cantidadCampanas);
        row.createCell(ci++).setCellValue(model.dotacionGpoTareas);
        row.createCell(ci++).setCellValue(model.cargoDesempenado);
        row.createCell(ci++).setCellValue(model.periodoDesde);
        row.createCell(ci++).setCellValue(model.periodoHasta);

        for (int i = 0; i < headers.length; i++) sheet.autoSizeColumn(i);
        sheet.createFreezePane(0, 1);

        String iosfaFile = (model.iosfa == null || model.iosfa.isBlank()) ? "ficha_datos" : model.iosfa.trim();
        String fileName = iosfaFile + ".xlsx";

        try (FileOutputStream fileOut = new FileOutputStream(fileName)) {
            workbook.write(fileOut);
        } finally {
            workbook.close();
        }

        return fileName;
    }

    // =========================
    // Helpers UI / formatos
    // =========================
    private String valueOf(JComboBox<String> cb) {
        Object v = cb == null ? null : cb.getSelectedItem();
        return v == null ? "" : v.toString().trim();
    }

    private JComboBox<String> newCombo() {
        JComboBox<String> cb = new JComboBox<>();
        cb.setEditable(false);
        cb.setPrototypeDisplayValue("XXXXXXXXXXXXXXXXXXXXXXXX");
        return cb;
    }

    private JTextField newNumericField(int columns) {
        JTextField field = new JTextField(columns);
        ((AbstractDocument) field.getDocument()).setDocumentFilter(new NumericDocumentFilter());
        return field;
    }

    private static class NumericDocumentFilter extends DocumentFilter {
        @Override
        public void insertString(FilterBypass fb, int offset, String string, AttributeSet attr) throws BadLocationException {
            if (string == null || string.matches("\\d*")) {
                super.insertString(fb, offset, string, attr);
            }
        }

        @Override
        public void replace(FilterBypass fb, int offset, int length, String text, AttributeSet attrs) throws BadLocationException {
            if (text == null || text.matches("\\d*")) {
                super.replace(fb, offset, length, text, attrs);
            }
        }
    }

    private JDateChooser newDatePicker() {
        JDateChooser dc = new JDateChooser();
        dc.setDateFormatString("dd/MM/yyyy");
        dc.setDate(null);
        if (dc.getDateEditor() instanceof JTextFieldDateEditor editor) {
            editor.setEditable(false);
        }
        return dc;
    }

    private String dateTextOf(JDateChooser dc) {
        if (dc == null) return "";
        Date d = dc.getDate();
        if (d == null) return "";
        SimpleDateFormat sdf = new SimpleDateFormat("dd/MM/yyyy");
        return sdf.format(d);
    }

    private void clearDatePicker(JDateChooser dc) {
        if (dc == null) return;
        dc.setDate(null);
    }

    private void addField(JPanel host, int row, String label, JComponent field) {
        GridBagConstraints c = new GridBagConstraints();
        c.gridx = 0; c.gridy = row * 2;
        c.weightx = 1.0;
        c.fill = GridBagConstraints.HORIZONTAL;
        c.anchor = GridBagConstraints.WEST;
        c.insets = new Insets(0, 0, 4, 0);

        JLabel lbl = new JLabel(label);
        lbl.setFont(lbl.getFont().deriveFont(Font.PLAIN, 12f));
        lbl.setForeground(FORM_LABEL_TEXT);
        host.add(lbl, c);

        c.gridx = 0; c.gridy = row * 2 + 1;
        c.insets = new Insets(0, 0, 12, 0);
        host.add(field, c);
    }

    private void styleButtonPrimary(JButton b, Color color) {
        b.setBackground(color);
        b.setForeground(Color.WHITE);
        b.setOpaque(true);
        b.setBorder(BorderFactory.createEmptyBorder(10, 18, 10, 18));
        b.setFocusPainted(false);
    }

    private void styleButtonSecondary(JButton b, Color color) {
        b.setBackground(color);
        b.setForeground(Color.WHITE);
        b.setOpaque(true);
        b.setBorder(BorderFactory.createEmptyBorder(10, 18, 10, 18));
        b.setFocusPainted(false);
    }

    private void styleButtonOutline(JButton b, Color color) {
        b.setBackground(Color.WHITE);
        b.setForeground(color);
        b.setOpaque(true);
        b.setBorder(BorderFactory.createLineBorder(color, 1));
        b.setFocusPainted(false);
    }

    private void styleTabbedPane(JTabbedPane tabs) {
        tabs.setOpaque(true);
        tabs.setFont(tabs.getFont().deriveFont(Font.BOLD, 13f));
        tabs.setBackground(new Color(0xEBF5FF));
        tabs.setForeground(FORM_HEADER_BG);
        tabs.setBorder(BorderFactory.createEmptyBorder(8, 8, 8, 8));
        tabs.setUI(new BasicTabbedPaneUI() {
            @Override
            protected void installDefaults() {
                super.installDefaults();
                tabAreaInsets = new Insets(0, 0, 0, 0);
                tabInsets = new Insets(8, 16, 8, 16);
                selectedTabPadInsets = new Insets(4, 12, 4, 12);
                contentBorderInsets = new Insets(12, 12, 12, 12);
            }

            @Override
            protected void paintTabArea(Graphics g, int tabPlacement, int selectedIndex) {
                Graphics2D g2 = (Graphics2D) g.create();
                g2.setColor(new Color(0xEBF5FF));
                g2.fillRect(0, 0, getWidth(), getHeight());
                g2.dispose();
                super.paintTabArea(g, tabPlacement, selectedIndex);
            }

            @Override
            protected void paintTabBackground(Graphics g, int tabPlacement, int tabIndex,
                                              int x, int y, int w, int h, boolean isSelected) {
                Graphics2D g2 = (Graphics2D) g.create();
                g2.setRenderingHint(RenderingHints.KEY_ANTIALIASING, RenderingHints.VALUE_ANTIALIAS_ON);
                g2.setColor(isSelected ? FORM_BUTTON_PRIMARY : new Color(0xDCE9FF));
                g2.fillRoundRect(x + 4, y + 4, Math.max(0, w - 8), Math.max(0, h - 6), 16, 16);
                g2.dispose();
            }

            @Override
            protected void paintTabBorder(Graphics g, int tabPlacement, int tabIndex,
                                          int x, int y, int w, int h, boolean isSelected) {
                // No border around tabs, rely on rounded backgrounds
            }

            @Override
            protected void paintText(Graphics g, int tabPlacement, Font font, java.awt.FontMetrics metrics,
                                     int tabIndex, String title, Rectangle textRect, boolean isSelected) {
                Graphics2D g2 = (Graphics2D) g.create();
                g2.setFont(font);
                g2.setRenderingHint(RenderingHints.KEY_TEXT_ANTIALIASING, RenderingHints.VALUE_TEXT_ANTIALIAS_ON);
                g2.setColor(isSelected ? Color.WHITE : FORM_HEADER_BG.darker());
                int textX = textRect.x;
                int textY = textRect.y + metrics.getAscent();
                g2.drawString(title, textX, textY);
                g2.dispose();
            }

            @Override
protected void paintContentBorder(Graphics g, int tabPlacement, int selectedIndex) {
    Graphics2D g2 = (Graphics2D) g.create();

    int tabAreaHeight = calculateTabAreaHeight(tabPlacement, runCount, maxTabHeight);
    int y = tabAreaHeight;

    g2.setColor(Color.WHITE);
    g2.fillRoundRect(0, y, getWidth(), getHeight() - y, 16, 16);

    g2.setColor(FORM_CARD_BORDER);
    g2.setStroke(new BasicStroke(1.2f));
    g2.drawRoundRect(0, y, getWidth() - 1, getHeight() - y - 1, 16, 16);

    g2.dispose();
}
        });
    }

    private JPanel createRoundedPanel(LayoutManager layout, Color background, int radius, Color borderColor, int borderWidth) {
        JPanel panel = new JPanel(layout) {
            @Override
            protected void paintComponent(Graphics g) {
                Graphics2D g2 = (Graphics2D) g.create();
                g2.setRenderingHint(RenderingHints.KEY_ANTIALIASING, RenderingHints.VALUE_ANTIALIAS_ON);

                int shadowOffset = 6;
                g2.setColor(new Color(0, 0, 0, 24));
                g2.fillRoundRect(shadowOffset, shadowOffset, getWidth() - shadowOffset * 2, getHeight() - shadowOffset * 2, radius, radius);

                g2.setColor(background);
                g2.fillRoundRect(0, 0, getWidth() - shadowOffset, getHeight() - shadowOffset, radius, radius);

                g2.setColor(borderColor);
                g2.setStroke(new BasicStroke(borderWidth));
                g2.drawRoundRect(borderWidth / 2, borderWidth / 2, getWidth() - shadowOffset - borderWidth, getHeight() - shadowOffset - borderWidth, radius, radius);

                g2.dispose();
                super.paintComponent(g);
            }

            @Override
            public boolean isOpaque() {
                return false;
            }
        };
        panel.setOpaque(false);
        return panel;
    }

    @SafeVarargs
    private final List<JTextField> forEachTextField(Supplier<JTextField>... f) {
        List<JTextField> out = new ArrayList<>();
        for (Supplier<JTextField> s : f) {
            try { out.add(s.get()); } catch (Exception ignored) {}
        }
        return out;
    }

    @SafeVarargs
    private final List<JDateChooser> forEachDatePicker(Supplier<JDateChooser>... f) {
        List<JDateChooser> out = new ArrayList<>();
        for (Supplier<JDateChooser> s : f) {
            try { out.add(s.get()); } catch (Exception ignored) {}
        }
        return out;
    }

    @SafeVarargs
    private final List<JComboBox<String>> forEachCombo(Supplier<JComboBox<String>>... f) {
        List<JComboBox<String>> out = new ArrayList<>();
        for (Supplier<JComboBox<String>> s : f) {
            try { out.add(s.get()); } catch (Exception ignored) {}
        }
        return out;
    }

    // =========================
    // Modelo
    // =========================
    static class FichaModel {
        String iosfa, dni, grado, apellido, nombres, lugarNac;

        String fechaNacimiento, fechaIngreso, aniosEnGrado, aniosEnEspecialidad, seDesempena, cumpleTurno;
        String escalafon, espBasica, espAvanzada, funcion, promedioTurnos, aptitudPsicofisicaFecha, fechaCondicionTiro, poseeAptoFisico;

        String cuil, cbu, emailInst, celular, usuarioGDE, rti, factorSanguineo, unidadRevista;
        String destinoInterno, cargo, destinoAnterior, deseaPermanecer;

        String domicilioCalle, domicilioNumero, localidad, cp, provincia;

        String estadoCivil, conyugeGrado, destino1, destino2;
        String observacionesEstadoCivil;

        // cónyuge base
        String conyugeApellido, conyugeNombre, conyugeFechaNacimiento, conyugeDNI;

        // cónyuge militar y condicionados
        String conyugeEsMilitar;
        String conyugeNroIdentificacion, conyugeGradoFAA, conyugeEspecialidad, conyugeDestino;

        String hijos, cantidadHijos, impedimentoTraslado;

        // Idiomas
        String idioma1, nivelIdioma1, fechaNivel1;
        String idioma2, nivelIdioma2, fechaNivel2;
        String idioma3, nivelIdioma3, fechaNivel3;
        String rindioSidiel;

        // Capacitaciones / docencia
        String capInstTitulo, capInstExpedidoPor, capInstFecha;
        String maxCapExtraTitulo, maxCapExtraExpedidoPor, maxCapExtraFecha;
        String actividadProfesor, asignaturaTemas, dictaActualmente, modalidad, tituloHabilitante;

        // Comisiones / campañas
        String comisionExteriorMotivo, comisionExteriorPaisCiudad, fechaInicioComision, fechaFinComision;
        String cumplioCampanasAntarticas, cantidadCampanas, dotacionGpoTareas, cargoDesempenado, periodoDesde, periodoHasta;
    }
}