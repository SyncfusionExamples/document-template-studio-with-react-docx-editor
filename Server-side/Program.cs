
namespace DocumentTemplateStudioService
{
    public class Program
    {
        internal static string webRootPath;
        
        public static void Main(string[] args)
        {
            var builder = WebApplication.CreateBuilder(args);

            builder.Services.AddControllersWithViews();

            builder.Services.AddCors(options =>
            {
                options.AddPolicy("AllowAllOrigins", builder =>
                {
                    builder.AllowAnyOrigin()
                    .AllowAnyMethod()
                    .AllowAnyHeader();
                });
            });

            var app = builder.Build();

            webRootPath = builder.Environment.WebRootPath;

            app.UseStaticFiles();

            app.UseRouting();

            app.UseCors("AllowAllOrigins");
            app.UseAuthorization();

            app.MapControllers();

            app.Run();
        }
    }
}
