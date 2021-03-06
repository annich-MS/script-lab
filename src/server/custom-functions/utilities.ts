import { parseMetadata } from './metadata.parser';
import { transformSnippetName } from '../core/snippet.helper';

export function getFunctionsAndMetadataForRegistration(
  snippets: ISnippet[]
): { visual: ICFVisualMetadata; functions: ICFFunctionMetadata[] } {
  const visualMetadata: ICFVisualSnippetMetadata[] = [];
  let metadata: ICFFunctionMetadata[] = [];

  snippets.filter(snippet => snippet.script && snippet.name).forEach(snippet => {
    let functions: ICFVisualFunctionMetadata[] = parseMetadata(
      snippet.script.content
    ) as ICFVisualFunctionMetadata[];

    functions = convertFunctionErrorsToSpace(functions);
    if (functions.length === 0) {
      return;
    } // no custom functions found
    const hasErrors = doesSnippetHaveErrors(functions);

    if (!hasErrors) {
      const namespace = transformSnippetName(snippet.name);
      const namespacedFunctions = functions.map(f => ({
        ...f,
        name: `${namespace}.${f.name}`,
      }));
      metadata = metadata.concat(...namespacedFunctions);
    }

    functions = functions.map(func => {
      const status: CustomFunctionsRegistrationStatus = hasErrors
        ? func.error
          ? 'error'
          : 'skipped'
        : 'good';

      func.parameters = func.parameters.map(p => ({
        ...p,
        prettyType: getPrettyType(p),
        status: getFunctionChildNodeStatus(func, status, p),
      }));

      return {
        ...func,
        paramString: paramStringExtractor(func), // todo, i think this can be removed
        status,
        result: {
          ...func.result,
          status: getFunctionChildNodeStatus(func, status, func.result),
        },
      };
    });

    // TODO:  why do we have code commented out?
    // const isTrusted = trustedSnippetManager.isSnippetTrusted(snippet.id, snippet.gist, snippet.gistOwnerId);
    // let status;
    // if (isTrusted) {
    let status: CustomFunctionsRegistrationStatus = hasErrors ? 'error' : 'good';
    // } else {
    //     status = CustomFunctionsRegistrationStatus.Untrusted;
    // }

    visualMetadata.push({
      name: transformSnippetName(snippet.name),
      error: hasErrors,
      status,
      functions,
    });
  });

  const functions = filterOutDuplicates(metadata);
  // const funcNames = functions.map(f => f.name);
  // const visual = { snippets: tagDuplicatesAsErrors(visualMetadata, funcNames) }; // todo see below
  const visual = { snippets: visualMetadata };

  return { visual, functions };
}

// helpers

function getFunctionChildNodeStatus(
  func: ICFVisualFunctionMetadata,
  funcStatus: CustomFunctionsRegistrationStatus,
  childNode: { error?: any }
): CustomFunctionsRegistrationStatus {
  return func.error ? (childNode.error ? 'error' : 'skipped') : funcStatus;
}

function getPrettyType(parameter) {
  if (parameter.error) {
    return '';
  }
  const dim = parameter.dimensionality === 'scalar' ? '' : '[][]';
  return `${parameter.type}${dim}`;
}

function paramStringExtractor(func) {
  if (func.error) {
    return undefined;
  }
  return func.parameters
    .map(p => {
      return `${p.name}: ${getPrettyType(p)}`;
    })
    .join(', ');
}

function doesSnippetHaveErrors(snippetMetadata) {
  return snippetMetadata.some(func => func.error);
}

// TODO REVISIT
// function tagDuplicatesAsErrors(
//   visualMetadata: ICFVisualSnippetMetadata[],
//   nonDuplicatedFunctionNames: string[]
// ): ICFVisualSnippetMetadata[] {
//   return visualMetadata.map(meta => {
//     let isError = meta.error;
//     meta.functions = meta.functions.map(func => {
//       if (!nonDuplicatedFunctionNames.includes(func.name) && !func.error) {
//         func.error =
//           ' - Duplicated function name. Must be unique across ALL snippets.';
//         func.status = CustomFunctionsRegistrationStatus.Error;
//         isError = true;
//       }
//       return func;
//     });
//     return {
//       ...meta,
//       error: isError,
//       status: isError
//         ? CustomFunctionsRegistrationStatus.Error
//         : CustomFunctionsRegistrationStatus.Good,
//     };
//   });
// }

/**
 * This function converts all the `true` errors on the functions to ' '. This is because we still want it
 * to have a truthy value, but not show anything in the UI, and this is the best way I could manage that at this time.
 * @param functions
 */
function convertFunctionErrorsToSpace(
  functions: ICFVisualFunctionMetadata[]
): ICFVisualFunctionMetadata[] {
  return functions.map(func => {
    if (func.error) {
      func.error = ' ';
    }
    return func;
  });
}

function filterOutDuplicates(functions: ICFFunctionMetadata[]): ICFFunctionMetadata[] {
  return functions.filter(func => {
    return functions.filter(f => f.name === func.name).length === 1;
  });
}
