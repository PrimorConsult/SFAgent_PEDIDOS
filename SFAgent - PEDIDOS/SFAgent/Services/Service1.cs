using Newtonsoft.Json;
using SFAgent.Salesforce;
using SFAgent.Sap;
using SFAgent.Utils;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.ServiceProcess;
using System.Threading;
using System.Threading.Tasks;

namespace SFAgent.Services
{
    public partial class Service1 : ServiceBase
    {
        private SalesforceAuth _auth;
        private SalesforceApi _api;
        private Timer _timer;

        public Service1()
        {
            InitializeComponent();
            _auth = new SalesforceAuth();
            _api = new SalesforceApi();
        }

        protected override void OnStart(string[] args)
        {
            if (!System.Diagnostics.EventLog.SourceExists("SFAgent"))
                System.Diagnostics.EventLog.CreateEventSource("SFAgent", "Application");

            Task.Run(async () =>
            {
                try
                {
                    await ProcessarPedidos();
                }
                catch (Exception ex)
                {
                    Logger.Log($"Erro inicial no OnStart (Pedidos): {ex.Message}");
                }
            });

            _timer = new Timer(async _ =>
            {
                try
                {
                    await ProcessarPedidos();
                }
                catch (Exception ex)
                {
                    Logger.Log($"Erro no Timer (Pedidos): {ex.Message}");
                }
            }, null, TimeSpan.FromMinutes(5), TimeSpan.FromMinutes(5));
        }

        protected override void OnStop()
        {
            _timer?.Dispose();
            Logger.Log("Serviço parado (Pedidos).");
        }

        // ---------- Helpers ----------
        private static bool IsDbNull(object v) => v == null || v == DBNull.Value;
        private static string S(object v) => IsDbNull(v) ? null : v.ToString();

        private static string AsDate(object dt)
        {
            if (IsDbNull(dt)) return null;
            var d = Convert.ToDateTime(dt, CultureInfo.InvariantCulture);
            return d.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture);
        }

        private static int? I(object v)
        {
            if (IsDbNull(v)) return null;
            try { return Convert.ToInt32(v, CultureInfo.InvariantCulture); } catch { return null; }
        }

        private static decimal? M2(object v)
        {
            if (IsDbNull(v)) return null;
            try
            {
                var d = Convert.ToDecimal(v, CultureInfo.InvariantCulture);
                return Math.Round(d, 2, MidpointRounding.AwayFromZero);
            }
            catch { return null; }
        }

        private static decimal? M3(object v)
        {
            if (IsDbNull(v)) return null;
            try
            {
                var d = Convert.ToDecimal(v, CultureInfo.InvariantCulture);
                return Math.Round(d, 3, MidpointRounding.AwayFromZero);
            }
            catch { return null; }
        }

        private static string Trunc(string s, int max)
            => string.IsNullOrEmpty(s) ? s : (s.Length <= max ? s : s.Substring(0, max));

        private static bool Yes(object v)
        {
            var s = S(v);
            return !string.IsNullOrWhiteSpace(s) &&
                   (s.Equals("Y", StringComparison.OrdinalIgnoreCase) ||
                    s == "1" ||
                    s.Equals("T", StringComparison.OrdinalIgnoreCase));
        }

        private static object Lookup(string ext) => string.IsNullOrWhiteSpace(ext) ? null : new { CA_IdExterno__c = ext };
        private static object Lookup(int? ext) => !ext.HasValue ? null : new { CA_IdExterno__c = ext.Value.ToString(CultureInfo.InvariantCulture) };

        private static string MapConsumidorFinal(string indFinal)
        {
            if (string.IsNullOrWhiteSpace(indFinal)) return null;
            return indFinal.Equals("Y", StringComparison.OrdinalIgnoreCase) ? "Sim" : "Não";
        }

        private static string MapTipoFrete(string tipo)
        {
            if (string.IsNullOrWhiteSpace(tipo)) return null;
            var t = tipo.Trim().ToUpperInvariant();
            if (t.StartsWith("CIF")) return "CIF";
            if (t.StartsWith("FOB")) return "FOB";
            return null; // evita erro de picklist
        }

        private static string MapMotivoCancelamento(string motivoRaw)
        {
            if (string.IsNullOrWhiteSpace(motivoRaw)) return null;
            var m = motivoRaw.Trim().ToLowerInvariant();
            if (m.Contains("condi") && m.Contains("pag")) return "Condição de Pagamento";
            if (m.Contains("crédit") || m.Contains("credito")) return "Crédito";
            if (m.Contains("estoq")) return "Estoque";
            if (m.Contains("prazo")) return "Prazo de Entrega";
            if (m.Contains("preç") || m.Contains("preco") || m.Contains("preço")) return "Preço";
            if (m.Contains("residual")) return "Quantidade Residual";
            if (m.Contains("tomada")) return "Tomada de Preço";
            if (m.Contains("troca")) return "Troca de Pedido";
            return null; // não força picklist
        }

        private static string MapStatus(string docStatus, string canceled, string printed)
        {
            if (Yes(canceled)) return "Cancelado"; // ORDR."CANCELED" = 'Y'
            if (string.Equals(docStatus, "C", StringComparison.OrdinalIgnoreCase)) return "Faturado";
            if (Yes(printed)) return "Impresso";
            if (string.Equals(docStatus, "O", StringComparison.OrdinalIgnoreCase)) return "Pendente";
            return "Rascunho";
        }


        private static string OnlyDigits(string s)
            => string.IsNullOrWhiteSpace(s) ? null : new string(s.Where(char.IsDigit).ToArray());
        // ---------- /Helpers ----------

        private async Task ProcessarPedidos()
        {
            try
            {
                var token = await _auth.GetValidToken();

                // Ajuste conforme ambiente
                var sap = new SapConnector("HANADB:30015", "SBO_ACOS_TESTE", "B1ADMIN", "S4P@2Q60_tm2");

                // NCLOB -> NVARCHAR via CTE (evita GROUP BY em LOB)
                var sql = @"
                            CALL SP_PEDIDOS_SF();
                          ";

                var rows = sap.ExecuteQuery(sql);

                int insertCount = 0;
                int updateCount = 0;
                int errorCount = 0;

                foreach (var r in rows)
                {
                    SalesforceApi.UpsertResult result = null;
                    string docNum = S(r["DocNum"]);
                    if (string.IsNullOrWhiteSpace(docNum))
                    {
                        Logger.Log("Pedido ignorado: DocNum vazio.");
                        continue;
                    }

                    try
                    {
                        // ----- Verificação de Existencia -----
                        bool pedidoExiste = await _api.ExistsPedido(token, docNum);

                        var statusSf = pedidoExiste
                            ? MapStatus(S(r["DocStatus"]), S(r["CANCELED"]), S(r["Printed"]))
                            : "Rascunho";


                        // ----- Cabeçalho -----
                        var cardCode = Trunc(S(r["CardCode"]), 50);
                        var cardName = Trunc(S(r["CardName"]), 200);
                        var name = Trunc(S(r["CardName"]), 80);
                        var cnpjCpf = Trunc(OnlyDigits(S(r["VATRegNum"])), 18); // normalizado
                        var filialId = I(r["BPLId"]);
                        var condPag = I(r["GroupNum"]);
                        var tipoFreteRaw = Trunc(S(r["TipoFrete"]), 10);
                        var rotaRaw = Trunc(S(r["Rota"]), 20);
                        var dtPedido = AsDate(r["DocDate"]);
                        var dtValidoAte = AsDate(r["DocDueDate"]);
                        var motivoCancel = MapMotivoCancelamento(S(r["MotivoCancel"]));
                        var obsFinais = Trunc(S(r["ObsFinais"]), 255);
                        var obsIniciais = Trunc(S(r["ObsIniciais"]), 255);
                        var totalAmount = M2(r["DocTotal"]);
                        var indFinal = MapConsumidorFinal(S(r["IndFinal"]));
                        var pnTri = Trunc(S(r["PnTriangular"]), 12);
                        var mainUsage = Trunc(S(r["MainUsage"]), 50);
                        var carrier = Trunc(S(r["Carrier"]), 50);
                        var formaPag = Trunc(S(r["PeyMethod"]), 40);
                        var deposito = Trunc(S(r["ToWhsCode"]), 8);

                        // ----- Endereço Entrega (S) -----
                        var entStreet = Trunc(S(r["StreetS"]), 254);
                        var entStreetNo = Trunc(S(r["StreetNoS"]), 20);
                        var entBlock = Trunc(S(r["BlockS"]), 100);   // Bairro
                        var entBuilding = Trunc(S(r["BuildingS"]), 255);// Complemento
                        var entCity = Trunc(S(r["CityS"]), 80);
                        var entZip = Trunc(S(r["ZipCodeS"]), 20);
                        var entState = Trunc(S(r["StateS"]), 2);
                        var entCountry = Trunc(string.IsNullOrWhiteSpace(S(r["CountryS"])) ? "BR" : S(r["CountryS"]), 2);

                        // ----- Endereço Cobrança (B) -----
                        var cobStreet = Trunc(S(r["StreetB"]), 254);
                        var cobStreetNo = Trunc(S(r["StreetNoB"]), 20);
                        var cobBlock = Trunc(S(r["BlockB"]), 100);
                        var cobBuilding = Trunc(S(r["BuildingB"]), 255);
                        var cobCity = Trunc(S(r["CityB"]), 80);
                        var cobZip = Trunc(S(r["ZipCodeB"]), 20);
                        var cobState = Trunc(S(r["StateB"]), 2);
                        var cobCountry = Trunc(string.IsNullOrWhiteSpace(S(r["CountryB"])) ? "BR" : S(r["CountryB"]), 2);

                        // ----- Agregados -----
                        var subTotal = M2(r["SubTotal"]);

                        var Pricebook2Id = S(r["Pricebook2Id"]);
                        var qtdItens = I(r["QtdItens"]);

                        // ----- Lookups por ExternalId -----
                        object contaLookup = Lookup(cardCode);
                        object filialLookup = Lookup(filialId);
                        object condPagLookup = Lookup(condPag);
                        object usoPrincipal2Lookup = Lookup(mainUsage);
                        object transportadoraLookup = Lookup(carrier);
                        object rotaLookup = Lookup(rotaRaw);
                        object formaPagLookup = Lookup(formaPag);
                        object depositoLookup = Lookup(deposito);
                        object pnTriLookup = Lookup(pnTri);

                        // ----- Mapeamentos finais -----
                        var tipoFrete = MapTipoFrete(tipoFreteRaw);
                        var transferencia = Yes(r["Transfered"]);

                        // ----- Body (anonymous object) -----
                        var body = new
                        {
                            // Identificação exibida
                            CA_NPedidoSAP__c = Trunc(docNum, 10),
                            Name = name,

                            // Cliente / Conta
                            CA_CodCliente__c = cardCode,
                            Account = contaLookup,

                            // Endereço entrega (compound)
                            CA_EnderecoEntrega__Street__s = entStreet,
                            CA_EnderecoEntrega__City__s = entCity,
                            CA_EnderecoEntrega__PostalCode__s = entZip,
                            CA_EnderecoEntrega__StateCode__s = entState,
                            CA_EnderecoEntrega__CountryCode__s = entCountry,
                            CA_BairroEntrega__c = entBlock,
                            CA_ComplementoEntrega__c = entBuilding,
                            CA_NumeroEntrega__c = entStreetNo,

                            // Endereço cobrança (compound)
                            CA_EnderecoCobranca__Street__s = cobStreet,
                            CA_EnderecoCobranca__City__s = cobCity,
                            CA_EnderecoCobranca__PostalCode__s = cobZip,
                            CA_EnderecoCobranca__StateCode__s = cobState,
                            CA_EnderecoCobranca__CountryCode__s = cobCountry,
                            CA_NumeroCobranca__c = cobStreetNo,
                            CA_BairroCobranca__c = cobBlock,
                            CA_ComplementoCobranca__c = cobBuilding,

                            // Dados fiscais/comerciais
                            CA_CPFCNPJ__c = cnpjCpf,
                            CA_Filial__r = filialLookup,
                            // (Comentado Devido a falta de IdExterno) CA_Deposito__r = depositoLookup,
                            CA_UsoPrincipal2__r = usoPrincipal2Lookup,
                            CA_CondicaodePagamento__r = condPagLookup,
                            // (Comentado Devido a falta de IdExterno) CA_FormaPagamento__r = formaPagLookup,
                            CA_TipoFrete__c = tipoFrete,
                            // (Comentado Devido a falta de IdExterno) CA_Rota__r = rotaLookup,
                            CA_Transportadora__r = transportadoraLookup,
                            CA_ConsumidorFinal__c = indFinal,
                            // (Comentado Devido a falta de IdExterno) CA_PnTriangular__r = pnTriLookup,

                            // Datas
                            EffectiveDate = dtPedido,
                            CA_ValidoAte__c = dtValidoAte,

                            // Observações / status
                            CA_MotivoCancelamento__c = motivoCancel,
                            CA_ObservacoesFinais__c = obsFinais,
                            CA_ObservacoesIniciais__c = obsIniciais,
                            Status = statusSf,

                            // Totais
                            CA_Subtotal__c = subTotal,
                            CA_QtdItens__c = qtdItens,
                            CA_TotalGeral__c = totalAmount,

                            // Integração
                            CA_Transferencia__c = transferencia,
                            CA_LimiteDisponivel__c = (decimal?)null,
                            CA_StatusIntegracao__c = "Integrado",
                            CA_RetornoIntegracao__c = "",
                            CA_AtualizacaoERP__c = DateTime.UtcNow.ToString("s") + "Z",

                            // Pricebook2 fixed Id.
                            Pricebook2Id = Pricebook2Id
                        };

                        // Upsert por ExternalId (apenas na URL)
                        var externalId = docNum;
                        result = await _api.UpsertPedido(token, externalId, body);

                        if (result?.Outcome != null)
                        {
                            var up = result.Outcome.ToUpperInvariant();
                            if (up == "POST" || up == "INSERT") insertCount++;
                            else if (up == "PATCH" || up == "UPDATE") updateCount++;
                        }

                        Logger.Log($"SF Pedidos {result?.Outcome} | METHOD={result?.Method} | ExternalId={externalId} | Status={result?.StatusCode}");
                    }
                    catch (Exception exItem)
                    {
                        errorCount++;
                        var rowJson = JsonConvert.SerializeObject(r);
                        Logger.Log(
                            $"ERRO Pedidos | ExternalId={docNum} | Msg={exItem.Message} | Row={rowJson}",
                            asError: true
                        );
                    }
                }

                Logger.Log($"Sync Pedidos finalizado. | Inseridos={insertCount} | Atualizados={updateCount} | Erros={errorCount} | Total Processados={rows.Count()}.");
            }
            catch (Exception ex)
            {
                Logger.Log($"Erro ao processar pedidos: {ex.Message}");
            }
        }
    }
}
