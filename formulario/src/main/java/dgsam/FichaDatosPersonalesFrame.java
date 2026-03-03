package dgsam;

import com.google.gson.Gson;
import com.google.gson.JsonObject;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import javax.swing.*;
import javax.swing.border.EmptyBorder;
import javax.swing.text.DateFormatter;
import javax.swing.text.DefaultFormatterFactory;

import java.awt.*;
import java.awt.event.ActionEvent;
import java.io.*;
import java.net.HttpURLConnection;
import java.net.URL;
import java.nio.charset.StandardCharsets;
import java.text.SimpleDateFormat;
import java.util.*;
import java.util.List;
import java.util.function.Supplier;

public class FichaDatosPersonalesFrame extends JFrame {

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

    // Datos personales (FECHAS: SPINNER)
    private JSpinner spFechaIngreso;
    private JTextField txtAniosEnGrado;
    private JTextField txtAniosEnEspecialidad;
    private JComboBox<String> cmbSeDesempena;
    private JComboBox<String> cmbCumpleTurno;
    private JTextField txtFuncion;
    private JTextField txtPromedioTurnos;
    private JSpinner spAptitudPsicofisicaFecha;
    private JSpinner spFechaCondicionTiro;
    private JComboBox<String> cmbPoseeAptoFisico;

    private JComboBox<String> cmbEspBasica;
    private JComboBox<String> cmbEspAvanzada;
    private JComboBox<String> cmbEscalafon;

    private JTextField txtEmailInst;
    private JSpinner spFechaNacimiento;
    private JTextField txtUsuarioGDE;
    private JTextField txtRTI;
    private JTextField txtFactorSanguineo;
    private JComboBox<String> cmbUnidadRevista;
    private JTextField txtDestinoInterno;
    private JTextField txtCargo;
    private JComboBox<String> cmbDestinoAnterior;
    private JComboBox<String> cmbDeseaPermanecer;

    // Idiomas
    private JTextField txtIdioma1;
    private JTextField txtNivelIdioma1;
    private JSpinner spFechaNivel1;
    private JTextField txtIdioma2;
    private JTextField txtNivelIdioma2;
    private JSpinner spFechaNivel2;
    private JTextField txtIdioma3;
    private JTextField txtNivelIdioma3;
    private JSpinner spFechaNivel3;
    private JComboBox<String> cmbRindioSidiel;

    // Capacitaciones y docencia
    private JTextField txtCapInstTitulo;
    private JTextField txtCapInstExpedidoPor;
    private JSpinner spCapInstFecha;
    private JTextField txtMaxCapExtraTitulo;
    private JTextField txtMaxCapExtraExpedidoPor;
    private JSpinner spMaxCapExtraFecha;
    private JComboBox<String> cmbActividadProfesor;
    private JTextField txtAsignaturaTemas;
    private JComboBox<String> cmbDictaActualmente;
    private JTextField txtModalidad;
    private JTextField txtTituloHabilitante;

    // Comisiones / campañas
    private JTextField txtComisionExteriorMotivo;
    private JTextField txtComisionExteriorPaisCiudad;
    private JSpinner spFechaInicioComision;
    private JSpinner spFechaFinComision;
    private JComboBox<String> cmbCumplioCampanasAntarticas;
    private JTextField txtCantidadCampanas;
    private JTextField txtDotacionGpoTareas;
    private JTextField txtCargoDesempenado;
    private JSpinner spPeriodoDesde;
    private JSpinner spPeriodoHasta;

    // ===== Domicilio =====
    private JTextField txtDomicilioCalle;
    private JTextField txtNumeroCalle;
    private JTextField txtLocalidad;
    private JButton btnGeolocalizar;

    // ===== Preferencias =====
    private JComboBox<String> cmbEstadoCivil;
    private JComboBox<String> cmbConyugeGrado;
    private JComboBox<String> cmbDestino1;
    private JComboBox<String> cmbDestino2;
    private JTextField txtObservEstadoCivil;
    private JTextField txtConyugeNroId;
    private JTextField txtConyugeGradoFAA;
    private JTextField txtConyugeApellido;
    private JTextField txtConyugeNombre;
    private JSpinner spConyugeFechaNacimiento;
    private JTextField txtConyugeDNI;
    private JTextField txtConyugeEspecialidad;
    private JTextField txtConyugeDestino;
    private JTextField txtHijos;
    private JTextField txtCantidadHijos;

    // TEXT AREA
    private JTextArea taImpedimentoTraslado;

    public FichaDatosPersonalesFrame(ListProvider listProvider) {
        super("Ficha de Datos Personales");
        this.listProvider = Objects.requireNonNull(listProvider, "listProvider");

        // asignar icono a la ventana (esquina superior izquierda junto al título)
        ImageIcon frameIcon = loadLogoImage();
        if (frameIcon != null) {
            setIconImage(frameIcon.getImage());
        }

        setDefaultCloseOperation(WindowConstants.DISPOSE_ON_CLOSE);
        setMinimumSize(new Dimension(1100, 700));
        setLocationRelativeTo(null);

        setContentPane(buildRoot());
        loadComboData();

        cmbEspBasica.setEnabled(false);
        cmbEspAvanzada.setEnabled(false);
        setupCascadeListeners();
    }

    private JComponent buildRoot() {
        JPanel root = new JPanel(new BorderLayout(0, 0));
        root.setBorder(new EmptyBorder(18, 18, 18, 18));
        root.setBackground(new Color(0xF7, 0xF7, 0xF7));

        JPanel card = new JPanel(new BorderLayout(0, 12));
        card.setBorder(BorderFactory.createCompoundBorder(
                BorderFactory.createLineBorder(new Color(0xE0, 0xE0, 0xE0), 1, true),
                new EmptyBorder(16, 16, 16, 16)
        ));
        card.setBackground(Color.WHITE);

        // Panel para título + logo
        JPanel headerPanel = new JPanel(new BorderLayout(5, 0));
        headerPanel.setBackground(Color.WHITE);
        
        JLabel title = new JLabel("Ficha de Datos Personales");
        title.setFont(title.getFont().deriveFont(Font.BOLD, 18f));
        headerPanel.add(title, BorderLayout.WEST);
        
        // Agregar logo pequeño al lado del título
        ImageIcon logo = loadLogoImage();
        if (logo != null) {
            JLabel logoLabel = new JLabel(logo);
            logoLabel.setMaximumSize(new Dimension(50, 50));
            logoLabel.setPreferredSize(new Dimension(50, 50));
            headerPanel.add(logoLabel, BorderLayout.EAST);
        }
        
        card.add(headerPanel, BorderLayout.NORTH);

        JTabbedPane tabs = new JTabbedPane();
        tabs.addTab("Datos personales", wrapScrollable(buildDatosPanel()));
        tabs.addTab("Domicilio", wrapScrollable(buildDomicilioPanel()));
        tabs.addTab("Preferencias", wrapScrollable(buildPreferenciasPanel()));
        card.add(tabs, BorderLayout.CENTER);

        card.add(buildFooterButtons(), BorderLayout.SOUTH);
        root.add(card, BorderLayout.CENTER);
        return root;
    }

    private ImageIcon loadLogoImage() {
        try {
            // Intentar cargar desde recursos
            InputStream is = getClass().getClassLoader().getResourceAsStream("logoDGSAM.jpg");
            if (is != null) {
                byte[] imageBytes = is.readAllBytes();
                is.close();
                ImageIcon icon = new ImageIcon(imageBytes);
                
                // Redimensionar a 50x50 píxeles
                Image scaledImage = icon.getImage().getScaledInstance(50, 50, Image.SCALE_SMOOTH);
                return new ImageIcon(scaledImage);
            }
        } catch (IOException e) {
            System.err.println("Error cargando logo: " + e.getMessage());
        }
        return null;
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

        spFechaIngreso = newDateSpinner();
        txtAniosEnGrado = new JTextField(6);
        txtAniosEnEspecialidad = new JTextField(6);
        cmbSeDesempena = newCombo();
        cmbCumpleTurno = newCombo();
        txtFuncion = new JTextField(18);
        txtPromedioTurnos = new JTextField(6);
        spAptitudPsicofisicaFecha = newDateSpinner();
        spFechaCondicionTiro = newDateSpinner();
        cmbPoseeAptoFisico = newCombo();
        spFechaNacimiento = newDateSpinner();

        cmbEspBasica = newCombo();
        cmbEspAvanzada = newCombo();
        cmbEscalafon = newCombo();

        txtCUIL = new JTextField(18);
        txtCBU = new JTextField(18);
        txtEmailInst = new JTextField(18);
        txtCelular = new JTextField(18);
        txtUsuarioGDE = new JTextField(18);
        txtRTI = new JTextField(12);
        txtFactorSanguineo = new JTextField(8);
        cmbUnidadRevista = newCombo();
        txtDestinoInterno = new JTextField(18);
        txtCargo = new JTextField(18);
        cmbDestinoAnterior = newCombo();
        cmbDeseaPermanecer = newCombo();

        // Idiomas
        txtIdioma1 = new JTextField(12);
        txtNivelIdioma1 = new JTextField(8);
        spFechaNivel1 = newDateSpinner();
        txtIdioma2 = new JTextField(12);
        txtNivelIdioma2 = new JTextField(8);
        spFechaNivel2 = newDateSpinner();
        txtIdioma3 = new JTextField(12);
        txtNivelIdioma3 = new JTextField(8);
        spFechaNivel3 = newDateSpinner();
        cmbRindioSidiel = newCombo();

        // Capacitaciones / docencia
        txtCapInstTitulo = new JTextField(18);
        txtCapInstExpedidoPor = new JTextField(18);
        spCapInstFecha = newDateSpinner();
        txtMaxCapExtraTitulo = new JTextField(18);
        txtMaxCapExtraExpedidoPor = new JTextField(18);
        spMaxCapExtraFecha = newDateSpinner();
        cmbActividadProfesor = newCombo();
        txtAsignaturaTemas = new JTextField(18);
        cmbDictaActualmente = newCombo();
        txtModalidad = new JTextField(12);
        txtTituloHabilitante = new JTextField(18);

        // Comisiones / campañas
        txtComisionExteriorMotivo = new JTextField(18);
        txtComisionExteriorPaisCiudad = new JTextField(18);
        spFechaInicioComision = newDateSpinner();
        spFechaFinComision = newDateSpinner();
        cmbCumplioCampanasAntarticas = newCombo();
        txtCantidadCampanas = new JTextField(6);
        txtDotacionGpoTareas = new JTextField(18);
        txtCargoDesempenado = new JTextField(18);
        spPeriodoDesde = newDateSpinner();
        spPeriodoHasta = newDateSpinner();

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
        addField(colRTop, rRTop++, "Factor sanguíneo", txtFactorSanguineo);
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
        addField(bottomFull, rb++, "Capacitación inst. - Título", txtCapInstTitulo);
        addField(bottomFull, rb++, "Capacitación inst. - Expedido por", txtCapInstExpedidoPor);
        addField(bottomFull, rb++, "Capacitación inst. - Fecha", spCapInstFecha);
        addField(bottomFull, rb++, "Máx. cap. extra - Título", txtMaxCapExtraTitulo);
        addField(bottomFull, rb++, "Máx. cap. extra - Expedido por", txtMaxCapExtraExpedidoPor);
        addField(bottomFull, rb++, "Máx. cap. extra - Fecha", spMaxCapExtraFecha);
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

        // Spacer final
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
        btnGeolocalizar = new JButton("Obtener ubicación aproximada");
        btnGeolocalizar.addActionListener(this::obtenerUbicacion);

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
        c.fill = GridBagConstraints.NONE;
        c.anchor = GridBagConstraints.CENTER;
        c.insets = new Insets(12, 0, 0, 0);
        panel.add(btnGeolocalizar, c);

        c.gridx = 0; c.gridy = 2;
        c.gridwidth = 2;
        c.weighty = 1.0;
        c.fill = GridBagConstraints.BOTH;
        panel.add(Box.createVerticalGlue(), c);

        return panel;
    }

    private JComponent buildPreferenciasPanel() {
        JPanel panel = new JPanel(new GridBagLayout());
        panel.setBackground(Color.WHITE);
        panel.setBorder(new EmptyBorder(8, 8, 8, 8));

        JPanel colL = new JPanel(new GridBagLayout());
        colL.setBackground(Color.WHITE);
        JPanel colR = new JPanel(new GridBagLayout());
        colR.setBackground(Color.WHITE);

        cmbEstadoCivil = newCombo();
        cmbConyugeGrado = newCombo();
        cmbDestino1 = newCombo();
        cmbDestino2 = newCombo();

        txtObservEstadoCivil = new JTextField(18);
        txtConyugeNroId = new JTextField(12);
        txtConyugeGradoFAA = new JTextField(12);
        txtConyugeApellido = new JTextField(18);
        txtConyugeNombre = new JTextField(18);
        spConyugeFechaNacimiento = newDateSpinner();
        txtConyugeDNI = new JTextField(12);
        txtConyugeEspecialidad = new JTextField(18);
        txtConyugeDestino = new JTextField(18);
        txtHijos = new JTextField(6);
        txtCantidadHijos = new JTextField(6);

        taImpedimentoTraslado = new JTextArea(6, 24);
        taImpedimentoTraslado.setLineWrap(true);
        taImpedimentoTraslado.setWrapStyleWord(true);
        JScrollPane spImpedimento = new JScrollPane(taImpedimentoTraslado);
        spImpedimento.setVerticalScrollBarPolicy(ScrollPaneConstants.VERTICAL_SCROLLBAR_AS_NEEDED);

        int rL = 0;
        addField(colL, rL++, "Estado civil", cmbEstadoCivil);
        addField(colL, rL++, "Cónyuge - Grado", cmbConyugeGrado);
        addField(colL, rL++, "Observaciones estado civil", txtObservEstadoCivil);
        addField(colL, rL++, "Cónyuge - Nro identificación", txtConyugeNroId);
        addField(colL, rL++, "Cónyuge - Grado FAA (si corresponde)", txtConyugeGradoFAA);
        addField(colL, rL++, "Cónyuge - Apellido", txtConyugeApellido);
        addField(colL, rL++, "Cónyuge - Nombre", txtConyugeNombre);
        addField(colL, rL++, "Cónyuge - Fecha nacimiento", spConyugeFechaNacimiento);
        addField(colL, rL++, "Cónyuge - DNI", txtConyugeDNI);
        addField(colL, rL++, "Cónyuge - Especialidad", txtConyugeEspecialidad);
        addField(colL, rL++, "Cónyuge - Destino", txtConyugeDestino);
        addField(colL, rL++, "Hijos (S/N)", txtHijos);
        addField(colL, rL++, "Cantidad hijos", txtCantidadHijos);

        int rR = 0;
        addField(colR, rR++, "Destino preferencia 1", cmbDestino1);
        addField(colR, rR++, "Destino preferencia 2", cmbDestino2);
        addField(colR, rR++, "Impedimento traslado (justificar)", spImpedimento);

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

    private JComponent buildFooterButtons() {
        JPanel footer = new JPanel(new BorderLayout());
        footer.setBackground(Color.WHITE);

        JPanel right = new JPanel(new FlowLayout(FlowLayout.RIGHT, 10, 0));
        right.setBackground(Color.WHITE);

        JButton btnLimpiar = new JButton("Limpiar");
        JButton btnGuardar = new JButton("Guardar");
        JButton btnCerrar = new JButton("Cerrar");

        styleButtonOutline(btnLimpiar, new Color(0x0D, 0x6E, 0xFD));
        styleButtonPrimary(btnGuardar, new Color(0x0D, 0x6E, 0xFD));
        styleButtonSecondary(btnCerrar, new Color(0x6C, 0x75, 0x7D));

        btnLimpiar.addActionListener(this::onLimpiar);
        btnGuardar.addActionListener(this::onGuardar);
        btnCerrar.addActionListener(e -> dispose());

        right.add(btnLimpiar);
        right.add(btnGuardar);
        right.add(btnCerrar);

        footer.add(right, BorderLayout.EAST);
        return footer;
    }

    private void loadComboData() {
        // Grado / provincias / etc (por listProvider)
        fillCombo(cmbGrado, listProvider.grados());
        fillCombo(cmbProvincia, listProvider.provincias());
        fillCombo(cmbEstadoCivil, listProvider.estadosCiviles());
        fillCombo(cmbConyugeGrado, listProvider.conyugeGrados());

        // ====== Escalafón/Especialidades DESDE EXCEL (resources) ======
        // TUS DATOS: columnas C, D, E y desde fila 8
        EspecialidadesData espData = readEspecialidadesDesdeExcelResource(
                "/ESPECIALIDADES DEL PERSONAL MILITAR SUBALTERNO (EN PROCESO).xlsx",
                0,   // sheet index
                1,   // B (0-based)
                2,   // C
                4,   // E
                7    // startRow: fila 8 (0-based)
        );

        if (espData != null && !espData.escalafones.isEmpty()) {
            fillCombo(cmbEscalafon, espData.escalafones);
            fillCombo(cmbEspBasica, espData.basicas);
            fillCombo(cmbEspAvanzada, espData.avanzadas);

            escalafon_especBasicaMap.clear();
            escalafon_especBasicaMap.putAll(espData.mapEscalafonToBasicas);

            especBasica_especAvanzadaMap.clear();
            especBasica_especAvanzadaMap.putAll(espData.mapBasicaToAvanzadas);
        } else {
            fillCombo(cmbEspBasica, listProvider.espBasica());
            fillCombo(cmbEspAvanzada, listProvider.espAvanzada());
            fillCombo(cmbEscalafon, listProvider.escalafon());

            escalafon_especBasicaMap.clear();
            especBasica_especAvanzadaMap.clear();
        }

        // Destinos desde CSV resources
        List<String> destinos = readDestinos();
        fillCombo(cmbDestino1, destinos);
        fillCombo(cmbDestino2, destinos);
        fillCombo(cmbUnidadRevista, destinos);
        fillCombo(cmbDestinoAnterior, destinos);

        List<String> siNo = Arrays.asList("", "Sí", "No");
        fillCombo(cmbSeDesempena, siNo);
        fillCombo(cmbCumpleTurno, siNo);
        fillCombo(cmbPoseeAptoFisico, siNo);
        fillCombo(cmbDeseaPermanecer, siNo);
        fillCombo(cmbRindioSidiel, siNo);
        fillCombo(cmbActividadProfesor, siNo);
        fillCombo(cmbDictaActualmente, siNo);
        fillCombo(cmbCumplioCampanasAntarticas, siNo);
    }

    private static class EspecialidadesData {
        final List<String> escalafones = new ArrayList<>();
        final List<String> basicas = new ArrayList<>();
        final List<String> avanzadas = new ArrayList<>();
        final Map<String, Set<String>> mapEscalafonToBasicas = new HashMap<>();
        final Map<String, Set<String>> mapBasicaToAvanzadas = new HashMap<>();
    }

    /**
     * Lee escalafón/básica/avanzada desde un Excel en resources.
     * - col indices 0-based: B=1, C=2, E=4
     * - startRow: 0-based (fila 8 => 7)
     *
     * Importante: hace "carry" de escalafón y básica cuando vienen en celdas
     * vacías (típico de Excel con celdas combinadas o valores agrupados).
     */
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

                    // carry: si vienen vacíos, repetir el último valor visto
                    String esc = !escRaw.isEmpty() ? escRaw : lastEsc;
                    String bas = !basRaw.isEmpty() ? basRaw : lastBas;
                    String avz = avzRaw; // avanzada generalmente no se "arrastra"

                    // si cambió escalafón explícitamente, actualizar y resetear básica arrastrada
                    if (!escRaw.isEmpty()) {
                        lastEsc = escRaw;
                        lastBas = ""; // reset por cambio de grupo
                    }
                    // si vino básica explícita, actualizar
                    if (!basRaw.isEmpty()) {
                        lastBas = basRaw;
                    }

                    // Saltar filas totalmente vacías
                    if (esc.isEmpty() && bas.isEmpty() && avz.isEmpty()) continue;

                    // Si el archivo tiene encabezado raro dentro, filtralo por texto típico
                    String escLow = esc.toLowerCase(Locale.ROOT);
                    String basLow = bas.toLowerCase(Locale.ROOT);
                    if (escLow.contains("escalaf") || basLow.contains("especialidad")) {
                        continue;
                    }

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
                // intentamos resolverla sin romper
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

    private void fillCombo(JComboBox<String> combo, List<String> items) {
        combo.removeAllItems();
        combo.addItem("");
        for (String s : items) combo.addItem(s);
        combo.setSelectedIndex(0);
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
                fillCombo(cmbEspBasica, new ArrayList<>());
                fillCombo(cmbEspAvanzada, new ArrayList<>());
            } else {
                cmbEspBasica.setEnabled(true);
                Set<String> basicas = escalafon_especBasicaMap.getOrDefault(escalafon, new TreeSet<>(String.CASE_INSENSITIVE_ORDER));
                fillCombo(cmbEspBasica, new ArrayList<>(basicas));
                cmbEspAvanzada.setEnabled(false);
                fillCombo(cmbEspAvanzada, new ArrayList<>());
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
                fillCombo(cmbEspAvanzada, new ArrayList<>());
            } else {
                cmbEspAvanzada.setEnabled(true);
                Set<String> avanzadas = especBasica_especAvanzadaMap.getOrDefault(basica, new TreeSet<>(String.CASE_INSENSITIVE_ORDER));
                fillCombo(cmbEspAvanzada, new ArrayList<>(avanzadas));
            }
        });
    }

    private void onLimpiar(ActionEvent e) {
        forEachTextField(
                () -> txtIOSFA, () -> txtDNI, () -> txtApellido, () -> txtNombres,
                () -> txtLugarNac, () -> txtCUIL, () -> txtCBU, () -> txtEmailInst, () -> txtCelular,
                () -> txtUsuarioGDE, () -> txtRTI, () -> txtFactorSanguineo,
                () -> txtDomicilioCalle, () -> txtNumeroCalle, () -> txtLocalidad, () -> txtCP,
                () -> txtObservEstadoCivil, () -> txtConyugeNroId, () -> txtConyugeGradoFAA, () -> txtConyugeApellido,
                () -> txtConyugeNombre, () -> txtConyugeDNI, () -> txtConyugeEspecialidad,
                () -> txtConyugeDestino, () -> txtHijos, () -> txtCantidadHijos,
                () -> txtFuncion, () -> txtPromedioTurnos,
                () -> txtIdioma1, () -> txtNivelIdioma1, () -> txtIdioma2, () -> txtNivelIdioma2, () -> txtIdioma3, () -> txtNivelIdioma3,
                () -> txtCapInstTitulo, () -> txtCapInstExpedidoPor, () -> txtMaxCapExtraTitulo, () -> txtMaxCapExtraExpedidoPor,
                () -> txtAsignaturaTemas, () -> txtModalidad, () -> txtTituloHabilitante,
                () -> txtComisionExteriorMotivo, () -> txtComisionExteriorPaisCiudad, () -> txtCantidadCampanas, () -> txtDotacionGpoTareas, () -> txtCargoDesempenado,
                () -> txtAniosEnGrado, () -> txtAniosEnEspecialidad,
                () -> txtDestinoInterno, () -> txtCargo
        ).forEach(tf -> tf.setText(""));

        if (taImpedimentoTraslado != null) taImpedimentoTraslado.setText("");

        forEachCombo(
                () -> cmbGrado, () -> cmbEspBasica, () -> cmbEspAvanzada, () -> cmbEscalafon,
                () -> cmbProvincia, () -> cmbEstadoCivil, () -> cmbConyugeGrado, () -> cmbDestino1, () -> cmbDestino2,
                () -> cmbSeDesempena, () -> cmbCumpleTurno, () -> cmbPoseeAptoFisico, () -> cmbDeseaPermanecer,
                () -> cmbRindioSidiel, () -> cmbActividadProfesor, () -> cmbDictaActualmente, () -> cmbCumplioCampanasAntarticas,
                () -> cmbUnidadRevista, () -> cmbDestinoAnterior
        ).forEach(cb -> cb.setSelectedIndex(0));

        forEachDateSpinner(
                () -> spFechaNacimiento, () -> spFechaIngreso, () -> spAptitudPsicofisicaFecha, () -> spFechaCondicionTiro,
                () -> spFechaNivel1, () -> spFechaNivel2, () -> spFechaNivel3, () -> spCapInstFecha, () -> spMaxCapExtraFecha,
                () -> spFechaInicioComision, () -> spFechaFinComision, () -> spPeriodoDesde, () -> spPeriodoHasta,
                () -> spConyugeFechaNacimiento
        ).forEach(this::clearSpinnerText);
    }

    private void onGuardar(ActionEvent e) {
        if (txtDNI.getText().trim().isEmpty()) {
            JOptionPane.showMessageDialog(this, "Debe ingresar DNI.", "Validación", JOptionPane.WARNING_MESSAGE);
            txtDNI.requestFocusInWindow();
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
        model.factorSanguineo = txtFactorSanguineo.getText().trim();
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
        model.conyugeGrado = valueOf(cmbConyugeGrado);
        model.destino1 = valueOf(cmbDestino1);
        model.destino2 = valueOf(cmbDestino2);

        model.observacionesEstadoCivil = txtObservEstadoCivil.getText().trim();
        model.conyugeNroIdentificacion = txtConyugeNroId.getText().trim();
        model.conyugeGradoFAA = txtConyugeGradoFAA.getText().trim();
        model.conyugeApellido = txtConyugeApellido.getText().trim();
        model.conyugeNombre = txtConyugeNombre.getText().trim();
        model.conyugeFechaNacimiento = dateTextOf(spConyugeFechaNacimiento);
        model.conyugeDNI = txtConyugeDNI.getText().trim();
        model.conyugeEspecialidad = txtConyugeEspecialidad.getText().trim();
        model.conyugeDestino = txtConyugeDestino.getText().trim();
        model.hijos = txtHijos.getText().trim();
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
            saveToExcel();
            JOptionPane.showMessageDialog(this, "Datos guardados en planilla", "Guardar", JOptionPane.INFORMATION_MESSAGE);
        } catch (IOException ex) {
            JOptionPane.showMessageDialog(this, "Error al guardar: " + ex.getMessage(), "Error", JOptionPane.ERROR_MESSAGE);
        }
    }

    private void saveToExcel() throws IOException {
        XSSFWorkbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("Ficha Datos");

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
                "Obs Estado Civil", "Conyuge Nro ID", "Conyuge Grado FAA", "Conyuge Apellido", "Conyuge Nombre",
                "Conyuge Fecha Nac", "Conyuge DNI", "Conyuge Especialidad", "Conyuge Destino",
                "Hijos", "Cantidad Hijos", "Impedimento Traslado",
                "Idioma 1", "Nivel 1", "Fecha nivel 1",
                "Idioma 2", "Nivel 2", "Fecha nivel 2",
                "Idioma 3", "Nivel 3", "Fecha nivel 3",
                "Rindió SIDIEL",
                "Cap Inst Título", "Cap Inst Expedido Por", "Cap Inst Fecha",
                "Max Cap Extra Título", "Max Cap Extra Expedido Por", "Max Cap Extra Fecha",
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
        row.createCell(ci++).setCellValue(model.cbu);
        row.createCell(ci++).setCellValue(model.emailInst);
        row.createCell(ci++).setCellValue(model.celular);
        row.createCell(ci++).setCellValue(model.usuarioGDE);
        row.createCell(ci++).setCellValue(model.rti);
        row.createCell(ci++).setCellValue(model.factorSanguineo);

        row.createCell(ci++).setCellValue(model.unidadRevista);
        row.createCell(ci++).setCellValue(model.destinoInterno);
        row.createCell(ci++).setCellValue(model.cargo);
        row.createCell(ci++).setCellValue(model.destinoAnterior);
        row.createCell(ci++).setCellValue(model.deseaPermanecer);

        row.createCell(ci++).setCellValue(model.domicilioCalle);
        row.createCell(ci++).setCellValue(model.domicilioNumero);
        row.createCell(ci++).setCellValue(model.localidad);
        row.createCell(ci++).setCellValue(model.cp);
        row.createCell(ci++).setCellValue(model.provincia);

        row.createCell(ci++).setCellValue(model.estadoCivil);
        row.createCell(ci++).setCellValue(model.conyugeGrado);
        row.createCell(ci++).setCellValue(model.destino1);
        row.createCell(ci++).setCellValue(model.destino2);

        row.createCell(ci++).setCellValue(model.observacionesEstadoCivil);
        row.createCell(ci++).setCellValue(model.conyugeNroIdentificacion);
        row.createCell(ci++).setCellValue(model.conyugeGradoFAA);
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

        String iosfaFile = model.iosfa == null || model.iosfa.isBlank()
        ? "ficha_datos"
        : model.iosfa.trim();

String fileName = iosfaFile + ".xlsx";

try (FileOutputStream fileOut = new FileOutputStream(fileName)) {
    workbook.write(fileOut);
}
        workbook.close();
    }

    // ===== Helpers =====

    private String valueOf(JComboBox<String> cb) {
        Object v = cb.getSelectedItem();
        return v == null ? "" : v.toString().trim();
    }

    private JComboBox<String> newCombo() {
        JComboBox<String> cb = new JComboBox<>();
        cb.setEditable(false);
        cb.setPrototypeDisplayValue("XXXXXXXXXXXXXXXXXXXXXXXX");
        return cb;
    }

    /**
     * Spinner de fecha (dd/MM/yyyy) con parseo estricto.
     * Si querés "vaciar" la fecha: podés borrar el texto manualmente, y al guardar se toma "".
     */
    private JSpinner newDateSpinner() {
        JSpinner sp = new JSpinner(new SpinnerDateModel(new Date(), null, null, Calendar.DAY_OF_MONTH));
        JSpinner.DateEditor ed = new JSpinner.DateEditor(sp, "dd/MM/yyyy");
        sp.setEditor(ed);

        JFormattedTextField tf = ed.getTextField();
        tf.setColumns(12);

        SimpleDateFormat sdf = new SimpleDateFormat("dd/MM/yyyy");
        sdf.setLenient(false);
        DateFormatter df = new DateFormatter(sdf);
        tf.setFormatterFactory(new DefaultFormatterFactory(df));
        tf.setFocusLostBehavior(JFormattedTextField.COMMIT_OR_REVERT);

        return sp;
    }

    private String dateTextOf(JSpinner sp) {
        if (sp == null) return "";
        if (!(sp.getEditor() instanceof JSpinner.DateEditor)) return "";
        JFormattedTextField tf = ((JSpinner.DateEditor) sp.getEditor()).getTextField();
        String t = tf.getText();
        return t == null ? "" : t.trim();
    }

    private void clearSpinnerText(JSpinner sp) {
        if (sp == null) return;
        if (sp.getEditor() instanceof JSpinner.DateEditor) {
            ((JSpinner.DateEditor) sp.getEditor()).getTextField().setText("");
        }
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
        host.add(lbl, c);

        c.gridx = 0; c.gridy = row * 2 + 1;
        c.insets = new Insets(0, 0, 12, 0);
        host.add(field, c);
    }

    private void styleButtonPrimary(JButton b, Color color) {
        b.setBackground(color);
        b.setForeground(Color.WHITE);
        b.setFocusPainted(false);
    }

    private void styleButtonSecondary(JButton b, Color color) {
        b.setBackground(color);
        b.setForeground(Color.WHITE);
        b.setFocusPainted(false);
    }

    private void styleButtonOutline(JButton b, Color color) {
        b.setBackground(Color.WHITE);
        b.setForeground(color);
        b.setFocusPainted(false);
    }

    @SafeVarargs
    private final List<JTextField> forEachTextField(Supplier<JTextField>... f) {
        List<JTextField> out = new ArrayList<>();
        for (Supplier<JTextField> s : f) out.add(s.get());
        return out;
    }

    @SafeVarargs
    private final List<JSpinner> forEachDateSpinner(Supplier<JSpinner>... f) {
        List<JSpinner> out = new ArrayList<>();
        for (Supplier<JSpinner> s : f) out.add(s.get());
        return out;
    }

    @SafeVarargs
    private final List<JComboBox<String>> forEachCombo(Supplier<JComboBox<String>>... f) {
        List<JComboBox<String>> out = new ArrayList<>();
        for (Supplier<JComboBox<String>> s : f) out.add(s.get());
        return out;
    }

    private void obtenerUbicacion(ActionEvent e) {
        try {
            URL url = new URL("http://ip-api.com/json");
            HttpURLConnection conn = (HttpURLConnection) url.openConnection();
            conn.setRequestMethod("GET");

            try (BufferedReader in = new BufferedReader(new InputStreamReader(conn.getInputStream(), StandardCharsets.UTF_8))) {
                StringBuilder content = new StringBuilder();
                String line;
                while ((line = in.readLine()) != null) content.append(line);

                Gson gson = new Gson();
                JsonObject json = gson.fromJson(content.toString(), JsonObject.class);

                if (json != null && json.has("status") && "success".equalsIgnoreCase(json.get("status").getAsString())) {
                    String city = json.has("city") ? json.get("city").getAsString() : "";
                    String region = json.has("regionName") ? json.get("regionName").getAsString() : "";
                    String zip = json.has("zip") ? json.get("zip").getAsString() : "";

                    txtLocalidad.setText(city);
                    txtCP.setText(zip);

                    for (int i = 0; i < cmbProvincia.getItemCount(); i++) {
                        if (cmbProvincia.getItemAt(i).equalsIgnoreCase(region)) {
                            cmbProvincia.setSelectedIndex(i);
                            break;
                        }
                    }

                    JOptionPane.showMessageDialog(this, "Ubicación aproximada obtenida.", "Geolocalización", JOptionPane.INFORMATION_MESSAGE);
                } else {
                    JOptionPane.showMessageDialog(this, "No se pudo obtener ubicación.", "Geolocalización", JOptionPane.WARNING_MESSAGE);
                }
            } finally {
                conn.disconnect();
            }
        } catch (Exception ex) {
            JOptionPane.showMessageDialog(this, "Error de conexión: " + ex.getMessage(), "Error", JOptionPane.ERROR_MESSAGE);
        }
    }

    // ===== Modelo =====
    static class FichaModel {
        String iosfa, dni, grado, apellido, nombres, lugarNac;

        String fechaNacimiento, fechaIngreso, aniosEnGrado, aniosEnEspecialidad, seDesempena, cumpleTurno;
        String escalafon, espBasica, espAvanzada, funcion, promedioTurnos, aptitudPsicofisicaFecha, fechaCondicionTiro, poseeAptoFisico;

        String cuil, cbu, emailInst, celular, usuarioGDE, rti, factorSanguineo, unidadRevista;
        String destinoInterno, cargo, destinoAnterior, deseaPermanecer;

        String domicilioCalle, domicilioNumero, localidad, cp, provincia;

        String estadoCivil, conyugeGrado, destino1, destino2;
        String observacionesEstadoCivil, conyugeNroIdentificacion, conyugeGradoFAA, conyugeApellido, conyugeNombre,
                conyugeFechaNacimiento, conyugeDNI, conyugeEspecialidad, conyugeDestino, hijos, cantidadHijos, impedimentoTraslado;

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