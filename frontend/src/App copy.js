import React, { useState, useRef } from 'react';
import {
  AppBar,
  Toolbar,
  Typography,
  Container,
  Box,
  Button,
  Card,
  CardContent,
  CardActions,
  Grid,
  Chip,
  Alert,
  LinearProgress,
  Stepper,
  Step,
  StepLabel,
  StepContent,
  Paper,
  List,
  ListItem,
  ListItemIcon,
  ListItemText,
  Divider,
  FormControlLabel,
  Switch,
  Accordion,
  AccordionSummary,
  AccordionDetails,
  Snackbar
} from '@mui/material';
import {
  CloudUpload,
  Description,
  Compare,
  Download,
  CheckCircle,
  Error,
  Warning,
  Info,
  ExpandMore,
  TableChart,
  Analytics,
  FilePresent,
  Settings
} from '@mui/icons-material';
import { createTheme, ThemeProvider } from '@mui/material/styles';

const theme = createTheme({
  palette: {
    primary: {
      main: '#1976d2',
    },
    secondary: {
      main: '#dc004e',
    },
  },
  typography: {
    h4: {
      fontWeight: 600,
    },
    h6: {
      fontWeight: 500,
    },
  },
});

const steps = [
  'Carregar Planilhas',
  'Configurar Opções',
  'Comparar Dados',
  'Baixar Resultado'
];

function App() {
  const [activeStep, setActiveStep] = useState(0);
  const [file1, setFile1] = useState(null);
  const [file2, setFile2] = useState(null);
  const [hasHeader1, setHasHeader1] = useState(true);
  const [hasHeader2, setHasHeader2] = useState(true);
  const [loading, setLoading] = useState(false);
  const [message, setMessage] = useState('');
  const [messageType, setMessageType] = useState('info');
  const [snackbarOpen, setSnackbarOpen] = useState(false);
  const [downloadUrl, setDownloadUrl] = useState(null);
  
  const file1Ref = useRef();
  const file2Ref = useRef();

  const handleFile1Change = (event) => {
    const selectedFile = event.target.files[0];
    if (selectedFile && selectedFile.name.endsWith('.xlsx')) {
      setFile1(selectedFile);
      if (activeStep === 0 && selectedFile && file2) {
        setActiveStep(1);
      }
    } else {
      showMessage('Por favor, selecione um arquivo .xlsx válido', 'error');
    }
  };

  const handleFile2Change = (event) => {
    const selectedFile = event.target.files[0];
    if (selectedFile && selectedFile.name.endsWith('.xlsx')) {
      setFile2(selectedFile);
      if (activeStep === 0 && file1 && selectedFile) {
        setActiveStep(1);
      }
    } else {
      showMessage('Por favor, selecione um arquivo .xlsx válido', 'error');
    }
  };

  const showMessage = (msg, type = 'info') => {
    setMessage(msg);
    setMessageType(type);
    setSnackbarOpen(true);
  };

  const handleCompare = async () => {
    if (!file1 || !file2) {
      showMessage('Por favor, selecione ambos os arquivos.', 'error');
      return;
    }

    setLoading(true);
    setActiveStep(2);

    const formData = new FormData();
    formData.append('file1', file1);
    formData.append('file2', file2);
    formData.append('has_header1', hasHeader1);
    formData.append('has_header2', hasHeader2);

    try {
      const response = await fetch('http://localhost:5000/upload', {
        method: 'POST',
        body: formData,
      });

      if (response.ok) {
        const blob = await response.blob();
        const url = window.URL.createObjectURL(blob);
        setDownloadUrl(url);
        setActiveStep(3);
        showMessage('Comparação realizada com sucesso! Arquivo pronto para download.', 'success');
      } else {
        const errorData = await response.json();
        showMessage(`Erro na comparação: ${errorData.error}`, 'error');
        setActiveStep(1);
      }
    } catch (error) {
      showMessage(`Erro de rede: ${error.message}`, 'error');
      setActiveStep(1);
    } finally {
      setLoading(false);
    }
  };

  const handleDownload = () => {
    if (downloadUrl) {
      const link = document.createElement('a');
      link.href = downloadUrl;
      link.download = 'comparacao_verbas.xlsx';
      document.body.appendChild(link);
      link.click();
      document.body.removeChild(link);
      showMessage('Download iniciado!', 'success');
    }
  };

  const resetProcess = () => {
    setActiveStep(0);
    setFile1(null);
    setFile2(null);
    setDownloadUrl(null);
    if (file1Ref.current) file1Ref.current.value = '';
    if (file2Ref.current) file2Ref.current.value = '';
  };

  return (
    <ThemeProvider theme={theme}>
      <Box sx={{ flexGrow: 1, minHeight: '100vh', backgroundColor: '#f5f5f5' }}>
        <AppBar position="static" elevation={0}>
          <Toolbar>
            <Analytics sx={{ mr: 2 }} />
            <Typography variant="h6" component="div" sx={{ flexGrow: 1 }}>
              Comparador de Verbas - Sistema Modernizado
            </Typography>
          </Toolbar>
        </AppBar>

        <Container maxWidth="lg" sx={{ mt: 4, mb: 4 }}>
          <Grid container spacing={3}>
            {/* Stepper */}
            <Grid item xs={12}>
              <Paper elevation={2} sx={{ p: 3 }}>
                <Stepper activeStep={activeStep} orientation="horizontal">
                  {steps.map((label, index) => (
                    <Step key={label}>
                      <StepLabel>{label}</StepLabel>
                    </Step>
                  ))}
                </Stepper>
              </Paper>
            </Grid>

            {/* Main Content */}
            <Grid item xs={12} md={8}>
              <Card elevation={2}>
                <CardContent>
                  {activeStep === 0 && (
                    <Box>
                      <Typography variant="h5" gutterBottom sx={{ display: 'flex', alignItems: 'center' }}>
                        <CloudUpload sx={{ mr: 1 }} />
                        Upload de Planilhas
                      </Typography>
                      <Typography variant="body2" color="text.secondary" sx={{ mb: 3 }}>
                        Selecione as duas planilhas Excel (.xlsx) que deseja comparar
                      </Typography>

                      <Grid container spacing={2}>
                        <Grid item xs={12} sm={6}>
                          <Card variant="outlined" sx={{ height: '100%' }}>
                            <CardContent>
                              <Typography variant="h6" gutterBottom>
                                <Description sx={{ mr: 1, verticalAlign: 'middle' }} />
                                Planilha 1 (Núcleo)
                              </Typography>
                              <Button
                                variant="contained"
                                component="label"
                                fullWidth
                                startIcon={<CloudUpload />}
                                sx={{ mb: 2 }}
                              >
                                Selecionar Arquivo
                                <input
                                  type="file"
                                  hidden
                                  ref={file1Ref}
                                  onChange={handleFile1Change}
                                  accept=".xlsx"
                                />
                              </Button>
                              {file1 && (
                                <Chip
                                  icon={<CheckCircle />}
                                  label={file1.name}
                                  color="success"
                                  variant="outlined"
                                  size="small"
                                />
                              )}
                            </CardContent>
                          </Card>
                        </Grid>

                        <Grid item xs={12} sm={6}>
                          <Card variant="outlined" sx={{ height: '100%' }}>
                            <CardContent>
                              <Typography variant="h6" gutterBottom>
                                <Description sx={{ mr: 1, verticalAlign: 'middle' }} />
                                Planilha 2 (Unidades)
                              </Typography>
                              <Button
                                variant="contained"
                                component="label"
                                fullWidth
                                startIcon={<CloudUpload />}
                                sx={{ mb: 2 }}
                              >
                                Selecionar Arquivo
                                <input
                                  type="file"
                                  hidden
                                  ref={file2Ref}
                                  onChange={handleFile2Change}
                                  accept=".xlsx"
                                />
                              </Button>
                              {file2 && (
                                <Chip
                                  icon={<CheckCircle />}
                                  label={file2.name}
                                  color="success"
                                  variant="outlined"
                                  size="small"
                                />
                              )}
                            </CardContent>
                          </Card>
                        </Grid>
                      </Grid>
                    </Box>
                  )}

                  {activeStep === 1 && (
                    <Box>
                      <Typography variant="h5" gutterBottom sx={{ display: 'flex', alignItems: 'center' }}>
                        <Settings sx={{ mr: 1 }} />
                        Configurações
                      </Typography>
                      <Typography variant="body2" color="text.secondary" sx={{ mb: 3 }}>
                        Configure as opções de processamento das planilhas
                      </Typography>

                      <Grid container spacing={2}>
                        <Grid item xs={12} sm={6}>
                          <Card variant="outlined">
                            <CardContent>
                              <Typography variant="h6" gutterBottom>
                                Planilha 1: {file1?.name}
                              </Typography>
                              <FormControlLabel
                                control={
                                  <Switch
                                    checked={hasHeader1}
                                    onChange={(e) => setHasHeader1(e.target.checked)}
                                  />
                                }
                                label="Contém cabeçalho"
                              />
                            </CardContent>
                          </Card>
                        </Grid>

                        <Grid item xs={12} sm={6}>
                          <Card variant="outlined">
                            <CardContent>
                              <Typography variant="h6" gutterBottom>
                                Planilha 2: {file2?.name}
                              </Typography>
                              <FormControlLabel
                                control={
                                  <Switch
                                    checked={hasHeader2}
                                    onChange={(e) => setHasHeader2(e.target.checked)}
                                  />
                                }
                                label="Contém cabeçalho"
                              />
                            </CardContent>
                          </Card>
                        </Grid>
                      </Grid>

                      <Box sx={{ mt: 3 }}>
                        <Button
                          variant="contained"
                          size="large"
                          startIcon={<Compare />}
                          onClick={handleCompare}
                          disabled={loading}
                          fullWidth
                        >
                          Comparar Planilhas Agora
                        </Button>
                      </Box>
                    </Box>
                  )}

                  {activeStep === 2 && (
                    <Box>
                      <Typography variant="h5" gutterBottom sx={{ display: 'flex', alignItems: 'center' }}>
                        <Compare sx={{ mr: 1 }} />
                        Processando Comparação
                      </Typography>
                      <Typography variant="body2" color="text.secondary" sx={{ mb: 3 }}>
                        Analisando as diferenças entre as planilhas...
                      </Typography>
                      <LinearProgress />
                    </Box>
                  )}

                  {activeStep === 3 && (
                    <Box>
                      <Typography variant="h5" gutterBottom sx={{ display: 'flex', alignItems: 'center' }}>
                        <CheckCircle sx={{ mr: 1 }} color="success" />
                        Comparação Concluída
                      </Typography>
                      <Typography variant="body2" color="text.secondary" sx={{ mb: 3 }}>
                        O arquivo Excel com as diferenças foi gerado com sucesso!
                      </Typography>

                      <Alert severity="success" sx={{ mb: 3 }}>
                        <Typography variant="body2">
                          O arquivo contém três abas: <strong>Diferenças</strong>, <strong>Estatísticas</strong> e <strong>Lado a Lado</strong>
                        </Typography>
                      </Alert>

                      <Box sx={{ display: 'flex', gap: 2, flexWrap: 'wrap' }}>
                        <Button
                          variant="contained"
                          size="large"
                          startIcon={<Download />}
                          onClick={handleDownload}
                          color="success"
                        >
                          Baixar Resultado
                        </Button>
                        <Button
                          variant="outlined"
                          size="large"
                          onClick={resetProcess}
                        >
                          Nova Comparação
                        </Button>
                      </Box>
                    </Box>
                  )}
                </CardContent>
              </Card>
            </Grid>

            {/* Sidebar */}
            <Grid item xs={12} md={4}>
              <Card elevation={2}>
                <CardContent>
                  <Typography variant="h6" gutterBottom sx={{ display: 'flex', alignItems: 'center' }}>
                    <Info sx={{ mr: 1 }} />
                    Instruções
                  </Typography>

                  <Accordion>
                    <AccordionSummary expandIcon={<ExpandMore />}>
                      <Typography variant="subtitle2">Formato dos Arquivos</Typography>
                    </AccordionSummary>
                    <AccordionDetails>
                      <List dense>
                        <ListItem>
                          <ListItemIcon>
                            <FilePresent fontSize="small" />
                          </ListItemIcon>
                          <ListItemText 
                            primary="Somente arquivos .xlsx"
                            secondary="Planilhas do Excel"
                          />
                        </ListItem>
                        <ListItem>
                          <ListItemIcon>
                            <TableChart fontSize="small" />
                          </ListItemIcon>
                          <ListItemText 
                            primary="3 colunas obrigatórias"
                            secondary="Matrícula, Verba, Valor"
                          />
                        </ListItem>
                      </List>
                    </AccordionDetails>
                  </Accordion>

                  <Accordion>
                    <AccordionSummary expandIcon={<ExpandMore />}>
                      <Typography variant="subtitle2">Processo de Comparação</Typography>
                    </AccordionSummary>
                    <AccordionDetails>
                      <List dense>
                        <ListItem>
                          <ListItemIcon>
                            <CloudUpload fontSize="small" />
                          </ListItemIcon>
                          <ListItemText primary="1. Carregar planilhas" />
                        </ListItem>
                        <ListItem>
                          <ListItemIcon>
                            <Settings fontSize="small" />
                          </ListItemIcon>
                          <ListItemText primary="2. Configurar opções" />
                        </ListItem>
                        <ListItem>
                          <ListItemIcon>
                            <Compare fontSize="small" />
                          </ListItemIcon>
                          <ListItemText primary="3. Executar comparação" />
                        </ListItem>
                        <ListItem>
                          <ListItemIcon>
                            <Download fontSize="small" />
                          </ListItemIcon>
                          <ListItemText primary="4. Baixar resultado" />
                        </ListItem>
                      </List>
                    </AccordionDetails>
                  </Accordion>

                  <Accordion>
                    <AccordionSummary expandIcon={<ExpandMore />}>
                      <Typography variant="subtitle2">Resultado</Typography>
                    </AccordionSummary>
                    <AccordionDetails>
                      <Typography variant="body2" color="text.secondary">
                        O arquivo Excel gerado contém três abas com análises detalhadas das diferenças encontradas entre as planilhas.
                      </Typography>
                    </AccordionDetails>
                  </Accordion>
                </CardContent>
              </Card>

              {/* Status Card */}
              <Card elevation={2} sx={{ mt: 2 }}>
                <CardContent>
                  <Typography variant="h6" gutterBottom>
                    Status do Sistema
                  </Typography>
                  <List dense>
                    <ListItem>
                      <ListItemIcon>
                        <CheckCircle color={file1 ? 'success' : 'disabled'} fontSize="small" />
                      </ListItemIcon>
                      <ListItemText primary="Planilha 1" secondary={file1 ? file1.name : 'Não carregada'} />
                    </ListItem>
                    <ListItem>
                      <ListItemIcon>
                        <CheckCircle color={file2 ? 'success' : 'disabled'} fontSize="small" />
                      </ListItemIcon>
                      <ListItemText primary="Planilha 2" secondary={file2 ? file2.name : 'Não carregada'} />
                    </ListItem>
                    <ListItem>
                      <ListItemIcon>
                        <CheckCircle color={downloadUrl ? 'success' : 'disabled'} fontSize="small" />
                      </ListItemIcon>
                      <ListItemText primary="Resultado" secondary={downloadUrl ? 'Pronto para download' : 'Aguardando comparação'} />
                    </ListItem>
                  </List>
                </CardContent>
              </Card>
            </Grid>
          </Grid>
        </Container>

        <Snackbar
          open={snackbarOpen}
          autoHideDuration={6000}
          onClose={() => setSnackbarOpen(false)}
        >
          <Alert 
            onClose={() => setSnackbarOpen(false)} 
            severity={messageType}
            sx={{ width: '100%' }}
          >
            {message}
          </Alert>
        </Snackbar>
      </Box>
    </ThemeProvider>
  );
}

export default App;

