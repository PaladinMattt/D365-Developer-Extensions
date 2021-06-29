﻿using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Client;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PluginDeployer.Spkl.Tasks
{
    public abstract class BaseTask
    {
        protected IOrganizationService _service;
        protected ITrace _trace;
        protected OrganizationServiceContext _context;

        public BaseTask(IOrganizationService service, ITrace trace)
        {
            _service = service;
            _trace = trace;
        }
        public BaseTask(OrganizationServiceContext context, ITrace trace)
        {
            _context = context;
            _trace = trace;

        }

        /// <summary>
        /// If set, a specific profile will be searched for in the spkl.json file - otherwise 'default' or none is used
        /// </summary>
        public string Profile { get; set; }
        public string Prefix { get; set; }
        public string Solution { get; set; }

        protected virtual void ExecuteInternal(string folder, OrganizationServiceContext ctx, bool backupFiles, string customClassRegex)
        {

        }
        
        public void Execute(string folder, bool backupFiles, string customClassRegex)
        {
            if (_context == null)
            {
                using (var ctx = new OrganizationServiceContext(_service))
                {
                    _context = ctx;
                    ctx.MergeOption = MergeOption.NoTracking;
                    ExecuteInternal(folder, ctx, backupFiles, customClassRegex);
                }
            }
            else
            {
                ExecuteInternal(folder, _context, backupFiles, customClassRegex);
            }
        }
    }
}